'use strict';
const Excel = require('exceljs');
const Joi = require('joi');
const Promise = require('bluebird');
const _ = require('lodash');

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
                    headerRow: Joi.number().min(1),
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
		// const rowOptions = this.options.rows;
        let isValid = true, lastMessage = '';

		// collecting sheet stats
		this.workbook.eachSheet((worksheet, sheetId) => {
            let name = worksheet.name;
            let sheetSchema = _.find(sheetOptions, {name: name});
            let result = this._checkSheet(worksheet, sheetSchema);
            if (result.isValid === false) {
                isValid = result.isValid;
                lastMessage = result.message;
            }
		});

		// after all checks, if boolean is false, then throw
		if (!isValid) {
			throw this._dataError(lastMessage);
		}
	}

    /** 
     * Checks a worksheet against its schema to make sure sheet is valid
     */
    _checkSheet(worksheet, sheetOptions) {
        let isValid = true, lastMessage = '';
        if (!sheetOptions || !sheetOptions.rows || _.isEmpty(sheetOptions.rows)) {
            isValid = false;
            lastMessage = 'Schema not found for sheet ' + worksheet.name;
        }
        else if (_.isNumber(sheetOptions.rows.headerRow) && sheetOptions.rows.allowedHeaders) {
            let rowOptions = sheetOptions.rows;
            let row = worksheet.getRow(rowOptions.headerRow);
            let headerNames = _.remove(row.values, null);
			const allowedHeaderNames = _.map(rowOptions.allowedHeaders, 'name');
			if (!_.chain(headerNames).difference(allowedHeaderNames).isEmpty().value()) {
				lastMessage = 'The row ' + rowOptions.headerRow + ' must contain headers. Only these header values are allowed: ' + allowedHeaderNames;
				isValid = false;
			}
		}

        return {
            isValid: isValid,
            message: lastMessage
        };
    }

    /** 
     * Error that is cause because of incorrect data is inputted to the class
     */
	_dataError(message) {
		return new Error(message);
	}

    /** 
     * Error that is caused by the class itself, and is not related to the
     * options provided by the user
     */
	_internalError(message) {
		return new Error('error while parsing excel file: ' + message);
	}

	getDefaultSheet() {
		if (this.options.sheets && this.options.sheets.default) {
			return this.workbook.getWorksheet();
		}
		else return null;
	}

	/**
	 * Returns a json version of the row data, based on the
	 * allowedHeaders of the given sheet. 
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
				let currentHeader = _.find(allowedHeaders, {name: header});
				let cellValue = cell.value;
                if (_.isObject(cell.value)) {
                    // If this is an object, then a formula has been applied
                    // We just take the result in that case
                    cellValue = cell.value.result;
                }
				result[currentHeader.key] = cellValue;
			}
		});
		return result;
	}

	/**
	 * Takes a callback and runs it on every row of the every sheet, one by one.
     * Order of the sheets is not guaranteed.
	 * This method provides each row in a json format based on the headers picked
	 * up from options
	 * callback params-
	 *  1. rowData, a json with key being the header field, picked up from options.row
	 *  2. rowNum, counting the headerRow
	 *  3. sheetKey, key of the sheet. If no key exists, the name is provided
	 *  The callback must return a promise
	 */
	eachRow(callback) {
		return this.afterRead.then(() => {
			let rowPromises = [];
            this.workbook.eachSheet((worksheet, sheetId) => {
                let sheetName = worksheet.name;
                let sheetOptions = _.find(this.options.sheets, {name: worksheet.name});
                let sheetKey = sheetOptions.key ? sheetOptions.key: sheetName;
                let headerRow = sheetOptions.rows.headerRow? sheetOptions.rows.headerRow: 1;
                let allowedHeaders = sheetOptions.rows.allowedHeaders;
                let headerRowValues = worksheet.getRow(headerRow).values;
                worksheet.eachRow((row, rowNum) => {
                    // ignoreing all the rows lesser than the headerRow
                    if (rowNum <= headerRow) {
                        return;
                    }
                    // processing the rest rows
                    let normalizedRowNum = rowNum - headerRow;
                    let rowData = this._getRowData(row, normalizedRowNum, allowedHeaders, headerRowValues);
                    rowPromises.push(callback(rowData, normalizedRowNum, sheetKey));
                });
            });
			return Promise.all(rowPromises);
		});
	}
}

module.exports = ExcelReader;