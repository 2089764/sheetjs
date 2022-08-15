# React

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags:

```html
<script src="xlsx.full.min.js"></script>
```

The library can also be imported directly from JSX code with:

```js
import { read, utils, writeFileXLSX } from 'xlsx';
```

This demo shows a simple React component transpiled in the browser using Babel
standalone library.  Since there is no standard React table model, this demo
settles on the array of arrays approach.


Other scripts in this demo show:
- server-rendered React component (with `next.js`)
- `react-native` deployment for iOS and android
- [`react-data-grid` reading, modifying, and writing files](modify/)

## How to run

Run `make react` to run the browser demo for React, or run `make next` to run
the server-rendered demo using `next.js`.

## Internal State

The simplest state representation is an array of arrays.  To avoid having the
table component depend on the library, the column labels are precomputed.  The
state in this demo is shaped like the following object:

```js
{
  cols: [{ name: "A", key: 0 }, { name: "B", key: 1 }, { name: "C", key: 2 }],
  data: [
    [ "id",    "name", "value" ],
    [    1, "sheetjs",    7262 ],
    [    2, "js-xlsx",    6969 ]
  ]
}
```

`sheet_to_json` and `aoa_to_sheet` utility functions can convert between arrays
of arrays and worksheets:

```js
/* convert from workbook to array of arrays */
var first_worksheet = workbook.Sheets[workbook.SheetNames[0]];
var data = XLSX.utils.sheet_to_json(first_worksheet, {header:1});

/* convert from array of arrays to workbook */
var worksheet = XLSX.utils.aoa_to_sheet(data);
var new_workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(new_workbook, worksheet, "SheetJS");
```

The column objects can be generated with the `encode_col` utility function:

```js
function make_cols(refstr/*:string*/) {
  var o = [];
  var range = XLSX.utils.decode_range(refstr);
  for(var i = 0; i <= range.e.c; ++i) {
    o.push({name: XLSX.utils.encode_col(i), key:i});
  }
  return o;
}
```

## React Native

[The new demo](https://docs.sheetjs.com/docs/demos/mobile#react-native) uses
up-to-date file I/O and file picker libraries.

## Server-Rendered React Components with Next.js

The demo reads from `public/sheetjs.xlsx`.  HTML output is generated using
`XLSX.utils.sheet_to_html` and inserted with `dangerouslySetInnerHTML`:

```jsx
export default function Index({html, type}) { return (
  // ...
  <div dangerouslySetInnerHTML={{ __html: html }} />
  // ...
); }
```

Next currently offers 3 general strategies for server-side data fetching:

#### "Server-Side Rendering" using `getServerSideProps`

`/getServerSideProps` reads the file on each request.  The first worksheet is
converted to HTML:

```js
export async function getServerSideProps() {
  const wb = XLSX.readFile(path);
  return { props: {
    html: utils.sheet_to_html(wb.Sheets[wb.SheetNames[0]])
  }};
}
```

#### "Static Site Generation" using `getStaticProps`

`/getServerSideProps` reads the file at build time.  The first worksheet is
converted to HTML:

```js
export async function getStaticProps() {
  const wb = XLSX.readFile(path);
  return { props: {
    html: utils.sheet_to_html(wb.Sheets[wb.SheetNames[0]])
  }};
}
```

#### "Static Site Generation with Dynamic Routes" using `getStaticPaths`

`/getStaticPaths` reads the file at build time and generates a list of sheets.

`/sheets/[id]` uses `getStaticPaths` to generate a path per sheet index:

```js
export async function getStaticPaths() {
  const wb = XLSX.readFile(path);
  return {
    paths: wb.SheetNames.map((name, idx) => ({ params: { id: idx.toString()  } })),
    fallback: false
  };
}
```

It also uses `getStaticProps` for the actual HTML generation:

```js
export async function getStaticProps(ctx) {
  const wb = XLSX.readFile(path);
  return { props: {
    html: utils.sheet_to_html(wb.Sheets[wb.SheetNames[ctx.params.id]]),
  }};
}
```

## Additional Notes

Some additional notes can be found in [`NOTES.md`](NOTES.md).

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
