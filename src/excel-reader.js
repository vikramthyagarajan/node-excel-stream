'use strict';
const Excel = require('exceljs');
const Joi = require('joi');
const Promise = require('bluebird');

class ExcelReader {
	/**
	 * Allowed options- {
     *      name: String,
     *      key: String,
	 *		sheets: [{
	 *			total: Number,
	 *			allowedNames: [String],
	 *			default: String
	 *		}]
	 * }
	 */
	constructor(dataStream, options) {
		this.stream = dataStream;
		this.options = options ? options: { sheets: {} };
		this.workbook = new Excel.Workbook();
		this.afterRead = this._read();
	}

	_read() {
		return this._validateOptions()
        .then(() => {
            return this.workbook.xlsx.read(this.stream);
        })
		.then((workbook) => {
			return this._checkWorkbook();
		})
		.then(() => {
			return this.workbook;
		});
	}

    /** 
     * Checks if options are of valid type and schema
     */
    _validateOptions() {
        let schema = Joi.object().keys({
            sheets: Joi.array().items(Joi.object().keys({
                name: Joi.string(),
                key: Joi.string(),
                allowedNames: Joi.array().items(Joi.string()),
                rows: Joi.object().keys({
                    headerRow: Joi.number(),
                    allowedHeaders: Joi.array().items(Joi.object().keys({
                        name: Joi.string(),
                        key: Joi.string()
                    }))
                })
            }))
        });

        return new Promise((resolve, reject) => {
            Joi.validate(this.options, schema, (err) => {
                if (err) {
                    reject(err);
                }
                else resolve();
            });
        });
    }

	_checkWorkbook() {
		// checks the workbook with the options specified
		// Used for error checking. Will give errors otherwise
		const sheetOptions = this.options.sheets;
		const rowOptions = this.options.rows;
		let total = 0, defaultSheet, names = [], headerNames = [];

		// collecting sheet stats
		this.workbook.eachSheet((worksheet, sheetId) => {
			total++;
			names.push(worksheet.name);
			if (rowOptions.headerRow) {
				let row = worksheet.getRow(rowOptions.headerRow);
				headerNames = _.remove(row.values, null);
			}
		});

		// checking sheet stats
		let boolean = true, lastMessage = '';
		if (sheetOptions.total && total !== sheetOptions.total) {
			lastMessage = 'Total number of sheets must be ', sheetOptions.total;
			boolean = false;
		}
		if (sheetOptions.allowedNames && !_.chain(names).difference(sheetOptions.allowedNames).isEmpty().value()) {
			lastMessage = 'Only these sheet names are allowed: ' + sheetOptions.allowedNames;
			boolean = false;
		}
		if (rowOptions && rowOptions.headerRow && rowOptions.allowedHeaders) {
			const allowedHeaderNames = _.map(rowOptions.allowedHeaders, 'text');
			if (!_.chain(headerNames).difference(allowedHeaderNames).isEmpty().value()) {
				lastMessage = 'The row ' + rowOptions.headerRow + ' must contain headers. Only these header values are allowed: ' + allowedHeaderNames;
				boolean = false;
			}
		}

		// after all checks, if boolean is false, then throw
		if (!boolean) {
			throw this._dataError(lastMessage);
		}
	}

	_dataError(message) {
		return new Boom.badData(message);
	}

	_internalError(message) {
		return new Boom.badImplementation('error while parsing excel file' + message);
	}

	getDefaultSheet() {
		if (this.options.sheets && this.options.sheets.default) {
			return this.workbook.getWorksheet();
		}
		else return null;
	}

	/**
	 * Returns a json version of the row data, based on the
	 * allowedHeaders of the default sheet. 
	 * Caveat: The allowedHeader must have an index option beforehand.
	 * This is to be calculated and set prior to calling this function.
	 */
	_getRowData(rowObject, rowNum, allowedHeaders, headerRowValues) {
		let result = {};
		rowObject.eachCell((cell, cellNo) => {
			// Finding the header value at this index
			if (!cell) {
				return;
			}
			let header = headerRowValues[cellNo];
			if (header) {
				let currentHeader = _.find(allowedHeaders, {text: header});
				let cellValue = cell.value;
				result[currentHeader.fieldName] = cellValue;
			}
		});
		return result;
	}

	/**
	 * Takes a callback and runs it on every row of the default sheet.
	 * Also, the default headers and header rows must be set, otherwise
	 * just get iterate the worksheets and do what you want.
	 * This method provides each row in a json format based on the headers picked
	 * up from options
	 * callback params-
	 *  1. rowData, a json with key being the header field, picked up from options.row
	 *  2. rowNum, counting the headerRow
	 *  The callback must return a promise
	 */
	eachRow(callback) {
		return this.afterRead.then(() => {
			let defaultSheet = this.getDefaultSheet();
			let rowPromises = [];
			// if (!defaultSheet) {
			// 	throw this._internalError('default sheet is not found');
			// }
			// if (!this.options.rows || !this.options.rows.headerRow || !this.options.rows.allowedHeaders) {
			// 	throw this._internalError('headerRow and allowedHeaders not set. Do not use this function. Use this.workbook itself');
			// }
			// let allowedHeaders = this.options.rows.allowedHeaders;
			// let headerRowValues = defaultSheet.getRow(this.options.rows.headerRow).values;
			// defaultSheet.eachRow((row, rowNum) => {
			// 	// ignoreing the headerRow
			// 	if (rowNum == this.options.rows.headerRow) {
			// 		return;
			// 	}
			// 	// processing the rest rows
			// 	let rowData = this._getRowData(row, rowNum, allowedHeaders, headerRowValues);
			// 	rowPromises.push(callback(rowData, rowNum));
			// })
			return Promise.all(rowPromises);
		});
	}
}

module.exports = ExcelReader;