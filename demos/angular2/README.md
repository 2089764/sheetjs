# Angular 2+

The ESM build can be imported directly from TS code with:

```typescript
import { read, utils, writeFileXLSX } from 'xlsx';
```

This demo uses an array of arrays (type `Array<Array<any>>`) as the core state.
The component template includes a file input element, a table that updates with
the data, and a button to export the data.

Other scripts in this demo show:
- `ionic` deployment for iOS, android, and browser
- `nativescript` deployment for iOS and android

## Array of Arrays

`Array<Array<any>>` neatly maps to a table with `ngFor`:

```html
<table class="sjs-table">
  <tr *ngFor="let row of data">
    <td *ngFor="let val of row">
      {{val}}
    </td>
  </tr>
</table>
```

The `aoa_to_sheet` utility function returns a worksheet.  Exporting is simple:

```typescript
/* generate worksheet */
const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

/* generate workbook and add the worksheet */
const wb: XLSX.WorkBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

/* save to file */
XLSX.writeFile(wb, 'SheetJS.xlsx');
```

`sheet_to_json` with the option `header:1` makes importing simple:

```typescript
/* <input type="file" (change)="onFileChange($event)" multiple="false" /> */
/* ... (within the component class definition) ... */
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const ab: ArrayBuffer = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(ab);

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
    };
    reader.readAsArrayBuffer(target.files[0]);
  }
```

## Switching between Angular versions

Modules that work with Angular 2 largely work as-is with Angular 4+.  Switching
between versions is mostly a matter of installing the correct version of the
core and associated modules.  This demo includes `package.json-angular#` files
for every major version of Angular up to 12.

To test a particular Angular version, overwrite `package.json`:

```bash
# switch to Angular 2
$ cp package.json-ng2 package.json
$ npm install
$ ng serve
```

Note: when running the demos, Angular 2 requires Node <= 14.  This is due to a
tooling issue with `ng` and does not affect browser use.

## XLSX Symbolic Link

In this tree, `node_modules/xlsx` is a link pointing back to the root.  This
enables testing the development version of the library.  In order to use this
demo in other applications, add the `xlsx` dependency:

```bash
$ npm install --save https://cdn.sheetjs.com/xlsx-latest/xlsx-latest.tgz
```

## SystemJS Configuration

The default angular-cli configuration requires no additional configuration.

Some deployments use the SystemJS loader, which does require configuration.
[SystemJS](https://docs.sheetjs.com/docs/demos/bundler#systemjs)
demo in the SheetJS CE docs describe the required settings.

## Ionic

<img src="screen.png" width="400px"/>

Reproducing the full project is a little bit tricky.  The included `ionic.sh`
script performs the necessary installation steps.

`Array<Array<any>>` neatly maps to a table with `ngFor`:

```html
<ion-grid>
  <ion-row *ngFor="let row of data">
    <ion-col *ngFor="let val of row">
      {{val}}
    </ion-col>
  </ion-row>
</ion-grid>
```


`@ionic-native/file` reads and writes files on devices. `readAsArrayBuffer`
returns `ArrayBuffer` objects suitable for `array` type, and `array` type can
be converted to blobs that can be exported with `writeFile`:

```typescript
/* read a workbook */
const ab: ArrayBuffer = await this.file.readAsArrayBuffer(url, filename);
const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'array'});

/* write a workbook */
const wbout: ArrayBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
let blob = new Blob([wbout], {type: 'application/octet-stream'});
this.file.writeFile(url, filename, blob, {replace: true});
```

## NativeScript

[The new demo](https://docs.sheetjs.com/docs/demos/mobile#nativescript)
is updated for NativeScript 8 and uses more idiomatic data patterns.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
