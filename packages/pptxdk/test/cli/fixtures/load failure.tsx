if (process.env.PPTXDK_TEST_LOAD_FAILURE === "1") {
  throw new Error("Failed to load author source module");
}

export default <div />;
