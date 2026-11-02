import { describe, expect, it } from "vitest";

import { greet } from "../src/index.ts";

describe("greet", () => {
  it("为有效姓名返回问候", () => {
    expect(greet("Ada")).toEqual({ message: "Hello, Ada" });
  });

  it("移除姓名两端的空白", () => {
    expect(greet("  Ada Lovelace  ")).toEqual({
      message: "Hello, Ada Lovelace",
    });
  });

  it("拒绝移除空白后为空的姓名", () => {
    expect(() => greet(" \t\n ")).toThrow();
  });
});
