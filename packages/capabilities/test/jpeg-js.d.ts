declare module "jpeg-js" {
  export interface RawImageData {
    readonly width: number;
    readonly height: number;
    readonly data: Uint8Array;
  }

  export function encode(
    imageData: RawImageData,
    quality?: number,
  ): { readonly data: Buffer; readonly width: number; readonly height: number };

  export function decode(
    jpegData: Uint8Array,
    options?: {
      readonly useTArray?: boolean;
      readonly colorTransform?: boolean;
      readonly formatAsRGBA?: boolean;
      readonly tolerantDecoding?: boolean;
      readonly maxResolutionInMP?: number;
      readonly maxMemoryUsageInMB?: number;
    },
  ): {
    readonly width: number;
    readonly height: number;
    readonly data: Uint8Array;
  };
}
