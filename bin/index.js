#!/usr/bin/env node

const co = require('co');
const getStdin = require('get-stdin');
const R = require('ramda');

co(function * () {
	try {
		const input = JSON.parse(yield getStdin());
		// if sheetOrder is specified, then reorder the keys
		// this is a problem since stringify/parse doesn't guarantee key order
		if (input.sheetOrder) {
			input.spreadsheet = R.reduce((acc, key) => {
				acc[key] = input.spreadsheet[key];
				return acc;
			}, {}, input.sheetOrder);
		}
		const xlsGen = require('../lib')(input.xlsGenOptions);
		const workbook = xlsGen.generate(input.spreadsheet, input.generateOptions);
		process.stdout.write(xlsGen.write(workbook, {bookType:'xlsx', type:'binary'}), 'binary');
	}
	catch (error) {
		console.error(error.stack);
		process.exit(1);
	}
});
