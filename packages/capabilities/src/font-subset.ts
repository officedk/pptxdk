import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

type HarfBuzzSubset = {
  readonly memory: WebAssembly.Memory;
  malloc(size: number): number;
  free(pointer: number): void;
  hb_blob_create(
    data: number,
    length: number,
    mode: number,
    userData: number,
    destroy: number,
  ): number;
  hb_blob_destroy(blob: number): void;
  hb_blob_get_data(blob: number, length: number): number;
  hb_blob_get_length(blob: number): number;
  hb_face_create(blob: number, index: number): number;
  hb_face_destroy(face: number): void;
  hb_face_reference_blob(face: number): number;
  hb_set_add(set: number, codePoint: number): void;
  hb_subset_input_create_or_fail(): number;
  hb_subset_input_destroy(input: number): void;
  hb_subset_input_unicode_set(input: number): number;
  hb_subset_or_fail(face: number, input: number): number;
};

const hbMemoryModeReadonly = 1;
const require = createRequire(import.meta.url);
let harfBuzzSubset: HarfBuzzSubset | undefined;

function loadHarfBuzzSubset(): HarfBuzzSubset {
  harfBuzzSubset ??= new WebAssembly.Instance(
    new WebAssembly.Module(
      readFileSync(require.resolve("harfbuzzjs/dist/harfbuzz-subset.wasm")),
    ),
  ).exports as unknown as HarfBuzzSubset;
  return harfBuzzSubset;
}

/** 依据 Unicode 字符集合保留真实 TrueType 字形及其布局闭包。 */
export function subsetFont(
  fontBytes: Uint8Array,
  characters: string,
): Uint8Array {
  if (fontBytes.byteLength === 0) throw new Error("字体数据为空");

  const hb = loadHarfBuzzSubset();
  const fontPointer = hb.malloc(fontBytes.byteLength);
  if (fontPointer === 0) throw new Error("HarfBuzz 无法分配字体内存");

  let sourceBlob = 0;
  let sourceFace = 0;
  let input = 0;
  let subsetFace = 0;
  let subsetBlob = 0;
  try {
    new Uint8Array(hb.memory.buffer, fontPointer, fontBytes.byteLength).set(
      fontBytes,
    );
    sourceBlob = hb.hb_blob_create(
      fontPointer,
      fontBytes.byteLength,
      hbMemoryModeReadonly,
      0,
      0,
    );
    if (sourceBlob === 0) throw new Error("HarfBuzz 无法读取字体数据");

    sourceFace = hb.hb_face_create(sourceBlob, 0);
    if (sourceFace === 0) throw new Error("HarfBuzz 无法读取字体字形表");
    hb.hb_blob_destroy(sourceBlob);
    sourceBlob = 0;

    input = hb.hb_subset_input_create_or_fail();
    if (input === 0) throw new Error("HarfBuzz 无法创建子集请求");
    const unicodeSet = hb.hb_subset_input_unicode_set(input);
    if (unicodeSet === 0) throw new Error("HarfBuzz 无法创建字符集合");
    for (const character of new Set(characters)) {
      const codePoint = character.codePointAt(0);
      if (codePoint !== undefined) hb.hb_set_add(unicodeSet, codePoint);
    }

    subsetFace = hb.hb_subset_or_fail(sourceFace, input);
    if (subsetFace === 0) throw new Error("HarfBuzz 无法生成字体子集");
    subsetBlob = hb.hb_face_reference_blob(subsetFace);
    if (subsetBlob === 0) throw new Error("HarfBuzz 未返回字体子集数据");
    const length = hb.hb_blob_get_length(subsetBlob);
    const pointer = hb.hb_blob_get_data(subsetBlob, 0);
    if (length === 0 || pointer === 0) {
      throw new Error("HarfBuzz 返回了空字体子集");
    }
    return new Uint8Array(hb.memory.buffer, pointer, length).slice();
  } finally {
    if (subsetBlob !== 0) hb.hb_blob_destroy(subsetBlob);
    if (subsetFace !== 0) hb.hb_face_destroy(subsetFace);
    if (input !== 0) hb.hb_subset_input_destroy(input);
    if (sourceFace !== 0) hb.hb_face_destroy(sourceFace);
    if (sourceBlob !== 0) hb.hb_blob_destroy(sourceBlob);
    hb.free(fontPointer);
  }
}
