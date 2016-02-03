var is = require('is_js');
var moment = require('moment');
var R = require('ramda');
var Ru = require('@panosoft/ramda-utils');
var XLSX = require('xlsx-style');

module.exports = options => {
	options = Ru.defaults({
		defaultDateFormat: 'mm/dd/yyyy', // sorry America is so stupid :-)
		isDateString: s => moment(s, moment.ISO_8601, true).isValid()
	},options || {});

	var excelDateToDate = excelDate => new Date((excelDate - 25569) * 86400 * 1000);
	var dateToExcelDate = date => (Date.parse(date) - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);

	var makeCell = (c, r, v, styles, defaultStyle) => {
		v = v || '';
		styles = styles || {};
		defaultStyle = defaultStyle || {};
		var cell = {v};
		var dataTypeStyle = {};
		if (is.json(v)) {
			cell = v;
			v = cell.v;
		}
		if (!cell.t) {
			if (is.number(v))
				cell.t = 'n';
			else if (is.boolean(v))
				cell.t = 'b';
			else if (is.date(v) || options.isDateString(v)) {
				v = new Date(v);
				cell.v = dateToExcelDate(v);
				cell.t = 'n';
				dataTypeStyle = {numFmt: options.defaultDateFormat};
			}
			else
				cell.t = 's';
		}
		// combine row and column styles (order of precedence from low to high: data type style, row style, column style, data style)
		var style = Ru.mergeAllR([defaultStyle, dataTypeStyle, styles.row || {}, (styles.columns || [])[c] || {}, cell.s || {}]);
		if (R.keys(style).length)
			cell.s = style;
		return cell;
	};

	var generate = (ss, defaultStyle) => {
		defaultStyle = defaultStyle || {};
		var workbook = {
			Sheets: {},
			SheetNames: []
		};
		var r, c, maxC;
		var nextRow = () => {++r; c = 0;};
		var nextColumn = () => {++c; maxC = Math.max(maxC, c);};
		R.forEach(sheetName => {
			r = 0;
			c = 0;
			maxC = 0;
			var worksheet = {};
			// header first
			var header = ss[sheetName].header;
			if (header) {
				R.forEach(headerValue => {
					worksheet[XLSX.utils.encode_cell({c,r})] = makeCell(c, r, headerValue, header.styles, defaultStyle.header);
					nextColumn();
				}, header.columns);
				nextRow();
			}
			// data next
			var data = ss[sheetName].data;
			R.forEach(row => {
				R.forEach(columnValue => {
					worksheet[XLSX.utils.encode_cell({c,r})] = makeCell(c, r, columnValue, data.styles, defaultStyle.data);
					nextColumn();
				}, row);
				nextRow();
			}, data.rows);
			worksheet['!ref'] = XLSX.utils.encode_range({s: {c: 0, r: 0}, e: {c: maxC, r}});
			workbook.SheetNames = R.append(sheetName, workbook.SheetNames);
			workbook.Sheets[sheetName] = worksheet;
		}, R.keys(ss));
		return workbook;
	};

	return {
		generate,
		write: XLSX.write,
		writeFile: XLSX.writeFile
	};
};
