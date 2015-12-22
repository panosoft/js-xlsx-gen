var xlsGen = require('../lib')({
	defaultDateFormat: 'yyyy-mm-dd' // that's more logical :-)
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
xlsGen.writeFile(workbook, 'workbook.xlsx');
//console.log(XLSX.write(workbook, {bookType:'xlsx', type:'binary'}));
