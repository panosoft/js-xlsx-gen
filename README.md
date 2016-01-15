# js-xlsx-gen

Simple declarative layer built on top of [`protobi/js-xlsx`](https://github.com/protobi/js-xlsx) which is a fork of [`SheetJS/js-xlsx`](https://github.com/SheetJS/js-xlsx) that adds styles.

[![npm version](https://img.shields.io/npm/v/js-xlsx-gen.svg)](https://www.npmjs.com/package/js-xlsx-gen)

## Installation

```sh
npm install js-xlsx-gen
```

## Usage

```js
var xlsGen = require('js-xlsx-gen')({
	defaultDateFormat: 'yyyy-mm-dd' // optional override of default date format of mm/dd/yyyy
});

// define some common styles
// see: https://github.com/protobi/js-xlsx/blob/master/tests/test-style.js for examples of styling AND https://github.com/SheetJS/ssf/blob/master/ssf.js for examples of numFmt
var commonStyles = {
	header: {
		alignment: {horizontal: 'center'},
		font: {bold: true}
	},
	positive: {
		numFmt: '$#,##0.00;$(#,##0.00)',
		font: {color: {rgb: '00ffff'}},
		fill: {fgColor: {rgb: 'ffffff'}}
	}
};
// next define the spreadsheet with each key as the sheet name
var spreadsheet = {
	sheetA: {
		header: {
			styles: {
				row: commonStyles.header,
				columns: {
					1: {font: {italic: true}}
				}
			},
			columns: [
				{v :'ColA', s: commonStyles.header},
				{v :'ColB', s: commonStyles.header}
			]
		},
		data: {
			rows: [
				[1, 2],
				[2, 3]
			]
		}
	},
	sheetB: {
		header: {
			columns: [
				'Column A',
				'Column B',
				'Column C',
				'Column D'
			]
		},
		data: {
			styles: {
				columns: {
					2: {numFmt: '$#,##0.00;$(#,##0.00)'},
					3: {font: {color: {rgb: '00ff00'}}}
				}
			},
			rows: [
				[
					new Date(),
					true,
					{v: new Date(), s: {numFmt: 'mm-dd-yyyy'}},
					20
				],
				[
					{v: -0.02, s: {numFmt: '$#,##0.00;$(#,##0.00)', font: {color: {rgb: 'ffffff'}}, fill: {fgColor: {rgb: 'ff0000'}}}}, // inline styling
					{v: 3, s: commonStyles.positive}, // inline styling to global styling
					1.23,
					40
				]
			]
		}
	}
};
// build workbook with OPTIONAL default cell styling
var workbook = xlsGen.generate(spreadsheet, {
	header: {font: { name: "Verdana", sz: 18, color: {rgb: "FF8800"}}, fill: {fgColor: {rgb: "aa00ff"}}},
	data: {font: { name: "Verdana", sz: 11, color: {rgb: "FF0000"}}, fill: {fgColor: {rgb: "ffaa00"}}}
});
// write out XLSX
xlsGen.writeFile(workbook, 'workbook.xlsx');
```
## API

- [`generate`](#generate)
- [`write`](#write)
- [`writeFile`](#writeFile)

---

<a name="generate"></a>
### generate ( spreadsheet, defaultStyle )
Generate a workbook from a spreadsheet definition with default cell styles.

__Arguments__

- `spreadsheet` - Spreadsheet definition
- `defaultStyle` - OPTIONAL styling for headers and data

__Example__

```js
var workbook = xlsGen.generate(spreadsheet, {
	header: {font: { name: "Verdana", sz: 18, color: {rgb: "FF8800"}}, fill: {fgColor: {rgb: "aa00ff"}}},
	data: {font: { name: "Verdana", sz: 11, color: {rgb: "FF0000"}}, fill: {fgColor: {rgb: "ffaa00"}}}
});
```

---

<a name="write"></a>
### write ( workbook, wopts )

Write workbook to Binary string. For more details see [`js-xlsx write()`](https://github.com/protobi/js-xlsx.git#writing-workbooks)

---

<a name="writeFile"></a>
### writeFile ( workbook, filePath )

Write workbook to Binary string. For more details see [`js-xlsx writeFile()`](https://github.com/protobi/js-xlsx.git#writing-workbooks)

## Spreadsheet Definition

The Spreadsheet Definition contains multiple sheets:

```js
var spreadsheet = {
	sheetA: {
	},
	sheetB: {		
	}
};
```

Each sheet has an optional `header` section and a `data` section:

```js
var spreadsheet = {
	sheetA: {
		header: {
		},
		data: {
		}
	}
};
```

The optional `header` section has an optional `style` section and `columns` section:

```js
var spreadsheet = {
	sheetA: {
		header: {
			styles: {
			},
			columns: [
			]
		}
	}
};
```

The `data` section has an optional `style` section and `data` section:

```js
var spreadsheet = {
	sheetA: {
		data: {
			styles: {
			},
			rows: [
			]
		}
	}
};
```

The `styles` section under the `header` and `data` sections contains an optional `row` section and an optional `columns` section:

```js
var spreadsheet = {
	sheetA: {
		header: {
			styles: {
				row: {
					alignment: {horizontal: 'center'},
					font: {bold: true}
				},
				columns {
					1: {font: {italic: true}}
				}
			},
			columns: [
			]
		},
		data: {
			styles: {
				row: {
					alignment: {horizontal: 'right'},
					font: {bold: true, italic: true, sz: 14}
				},
				columns {
					1: {
						numFmt: '$#,##0.00;$(#,##0.00)',
						font: {color: {rgb: '00ffff'}},
						fill: {fgColor: {rgb: 'ffffff'}}
					}
				}
			},
			rows: [
			]
		}
	}
};
```

Unfortunately, these styles aren't documented well, but examples can be found in the [`style test code of protobi/js-xlsx`](https://github.com/protobi/js-xlsx/blob/master/tests/test-style.js).

The `columns` array of the `header` section contains cell data:

```js
var spreadsheet = {
	sheetA: {
		header: {
			columns: [
				{v :'ColA', s: commonStyles.header},
				'ColB'
			]
		},
		data: {
		}
	}
};
```

Notice that the `columns` array contains either objects that conform to [`Cell Objects in protobi/js-xlsx`](https://github.com/protobi/js-xlsx#cell-object) OR raw data, i.e. number, date, string, boolean.

The `rows` array of the `data` section contains row data:

```js
var spreadsheet = {
	sheetA: {
		data: {
			rows: [
				[
					new Date(),
					true,
					{v: new Date(), s: {numFmt: 'mm-dd-yyyy'}},
					20
				],
				[
					{v: -0.02, s: {numFmt: '$#,##0.00;$(#,##0.00)', font: {color: {rgb: 'ffffff'}}, fill: {fgColor: {rgb: 'ff0000'}}}},
					{v: 3, s: commonStyles.positive},
					1.23,
					40
				]
			]
		}
	}
};
```

Notice that the `rows` array contains an *array for each row* which contains either objects that conform to [`Cell Objects in protobi/js-xlsx`](https://github.com/protobi/js-xlsx#cell-object) OR raw data, i.e. number, date, string, boolean.
