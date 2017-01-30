'use strict';
const expect = require('chai').expect;
const Excel = require('exceljs');
const ExcelReader = require('../index').ExcelReader;
const fs = require('fs');

let testWorkbooks = {};
describe('Excel Reader', () => {
    before(() => {
        testWorkbooks.singleSheetFirstRowHeader = fs.createReadStream('./util/excels/1sheet-1header.xlsx');
    });
    describe('Sheets', () => {
        it('should error if different number of Sheets', (done) => {
            let workbook = testWorkbooks.singleSheetFirstRowHeader;
            let reader = new ExcelReader(testWorkbooks.singleSheetFirstRowHeader, {
                sheets: [{
                    name: 'Data',
                }, {
                    name: 'Test'
                }]
            })
            reader.eachRow()
            .then(() => {
                done('Reader must exit with an error');
            })
            .catch((err) => {
                expect(err).to.exist();
            });
        });
    });

    describe('Allowed Sheet Names', () => {
        it('should be an array', () => {
        });

        it('should only allow selected sheet names', () => {
        });

        it('should throw error if invalid sheet name is in excel', () => {
        });

        it('should allow any name if allowedNames is null', () => {
        });
    });

    describe('Rows', () => {
        it('should error if header row does not exist in excel', () => {
        });
        
        it('should only allow specified header rows', () => {
        });

        it('should return each row data based on header row keys', () => {
        });
    });
});