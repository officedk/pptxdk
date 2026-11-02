if (process.env.REACT_PPTX_TEST_LOAD_FAILURE === "1") {
  throw new Error("作者源码模块加载失败");
}

export default <div />;
