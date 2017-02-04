'use strict';
const expect = require('chai').expect;
const Excel = require('exceljs');
const ExcelReader = require('../index').ExcelReader;
const fs = require('fs');
const rm = require('rimraf');

let testWorkbooks = {};
let tempDirectoryPath = __dirname + '/temp';
describe('Excel Writer', () => {
    before(() => {
        // before the tests here, we must create a temp directory, and write the excels there
        // When checking the excels is required, then a read stream is returned for that file
        // using these functions-
        testWorkbooks.singleSheetFirstRowHeader = () => fs.createReadStream(__dirname + '/util/excels/1sheet-1header.xlsx');
        testWorkbooks.singleSheetNRowHeader = () => fs.createReadStream(__dirname + '/util/excels/1sheet-nheader.xlsx');
        testWorkbooks.multiSheetNRowHeader = () => fs.createReadStream(__dirname + '/util/excels/2sheet-nheader.xlsx');

        return fs.mkdir(tempDirectoryPath);
    });

    describe('Something', () => {
    });

    after(() => {
        return new Promise((resolve, reject) => {
            return rm(tempDirectoryPath, (err) => {
                if (err)
                    reject(err);
                else 
                    resolve();
            });
        });
    });
});