'use strict';
const _ = require('lodash');
const Excel = require('exceljs');
const Stream = require('stream');

class ExcelWriter {
	/**
	 * options currently contains all are required. Arrays can be empty-
	 * sheets: [{
	 *		name: String, //Name shwos up
	 *		key: String, // Identifier to push data to this sheet
	 *		headers: [{
	 *			text: String,
	 *			key: String,
	 *			default: any
	 *		}]
	 * }]
	 */
	constructor(options) {
		this.stream = Stream.Transform({
		  write: function(chunk, encoding, next) {
			this.push(chunk);
			next();
		  }
		});
		this.options = options ? options: {sheets: []};
		if (this.options.debug) {
			console.log('creating workbook');
		}
		this.workbook = new Excel.stream.xlsx.WorkbookWriter({
			stream: this.stream
		});
		this.worksheets = {};
		this.rowCount = 0;
		this.sheetRowCount = {};
		this.afterInit = this._createWorkbookLayout();
	}

	/**
	 * Function that starts writing the sheets etc given in the options
	 * and keeps the workbook reader from writing data
	 */
	_createWorkbookLayout() {
        let error;
		this.options.sheets.map((sheetSpec) => {
            if (!sheetSpec.key) {
                error = 'No key specified for sheet: ' + sheetSpec.name;
            }
			if (this.options.debug) {
				console.log('creating sheet', sheetSpec.name);
			}
			let sheet = this.workbook.addWorksheet(sheetSpec.name);
			let headerNames = _.map(sheetSpec.headers, 'name');
			sheet.addRow(headerNames).commit();
			this.worksheets[sheetSpec.key] = sheet;
		});
		return new Promise((resolve, reject) => {
            if (error) {
                reject(this._dataError(error));
            }
            else resolve();
        });
	}

    _dataError(message) {
        return new Error(message);
    }

	addData(sheetKey, dataObj) {
		return this.afterInit.then(() => {
			let sheet = this.worksheets[sheetKey];
			let sheetSpec = _.find(this.options.sheets, {key: sheetKey});
			let rowData = [];
			if (!sheet)
				throw this._dataError('No such sheet key: ' + sheetKey);
			if (!sheetSpec)
				throw this._dataError('No such sheet key ' + sheetKey + ' in the spec. Check options');
			sheetSpec.headers.map((header, index) => {
				let defaultValue = '';
				if (header.default !== undefined && header.default !== null) {
					defaultValue = header.default;
				}
				if (dataObj[header.key]) {
					rowData[index] = dataObj[header.key];
				}
				else rowData[index] = defaultValue;
			});
			return Promise.resolve({
				sheet: sheet,
				data: rowData
			});
		})
		.then((rowObj) => {
			let sheet = rowObj.sheet;
			let data = rowObj.data;
			this.rowCount++;
			this.sheetRowCount[sheetKey] = this.sheetRowCount[sheetKey]? this.sheetRowCount[sheetKey] + 1: 1;
			return sheet.addRow(data).commit();
		});
	}

	save() {
		return this.afterInit.then(() => {
			_.map(this.worksheets, (worksheet, worksheetName) => {
				worksheet.commit();
			});
			if (this.options.debug) {
				console.log('written ' + this.rowCount + ' rows in total');
				console.log('written rows in each sheet -', this.sheetRowCount);
				console.log('commiting and closing the workbook');
			}
			return this.workbook.commit()
            .then(() => {
                return this.stream;
            });
		});
	}
}

module.exports = ExcelWriter;