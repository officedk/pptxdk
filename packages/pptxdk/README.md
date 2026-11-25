# @officedk/pptxdk

Create editable PowerPoint presentations with JavaScript, TypeScript or React. The package includes the SDK and the `pptxdk` command for building, previewing and generating source from presentations.

## Build your first presentation

Install the package in a project running Node.js 26.13 or a later 26.x release.

```sh
npm install @officedk/pptxdk
```

Save the following as `slides.mjs`.

```js
import { deck, shape, slide } from "@officedk/pptxdk/vanilla";

export default deck({ width: 960, height: 540 }).append(
  slide({ id: "intro" }).append(
    shape({
      id: "title",
      x: 64,
      y: 80,
      width: 600,
      height: 96,
      text: "My first presentation",
    }),
  ),
);
```

Run the CLI from the project directory to write `slides.pptx`.

```sh
npx pptxdk build slides.mjs
```

To use the SDK from another ES module, pass the same document to `build`.

```js
import { writeFile } from "node:fs/promises";

import { build } from "@officedk/pptxdk";

import slides from "./slides.mjs";

const { pptx } = await build(slides);
await writeFile("slides.pptx", pptx);
```
