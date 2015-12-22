#!/usr/bin/env node

const co = require('co');
const getStdin = require('get-stdin');


co(function * () {
	try {
		const input = JSON.parse(yield getStdin());
		console.error(input.xlsGenOptions)
		const xlsGen = require('../lib')(input.xlsGenOptions);
		const workbook = xlsGen.generate(input.spreadsheet, input.generateOptions);
		process.stdout.write(xlsGen.write(workbook, {bookType:'xlsx', type:'binary'}), 'binary');
	}
	catch (error) {
		console.error(error.stack);
		process.exit(1);
	}
});
