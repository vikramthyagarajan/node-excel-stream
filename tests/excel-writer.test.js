'use strict';
const expect = require('chai').expect;
const Excel = require('exceljs');
const ExcelWriter = require('../index').ExcelWriter;
const fs = require('fs');
const rm = require('rimraf');

let readWorkbooks = {}, writeWorkbooks = {};
let tempDirectoryPath = __dirname + '/temp';
describe('Excel Writer', () => {
    before(() => {
        // before the tests here, we must create a temp directory, and write the excels there
        // When checking the excels is required, then a read stream is returned for that file
        // using these functions-
        readWorkbooks.onlyHeaders = () => fs.createReadStream(tempDirectoryPath + '/only-headers.xlsx');
        readWorkbooks.defaults = () => fs.createReadStream(tempDirectoryPath + '/defaults.xlsx');
        // readWorkbooks.multiSheetNRowHeader = () => fs.createReadStream(tempDirectoryPath + '/2sheet-nheader.xlsx');

        writeWorkbooks.onlyHeaders = () => fs.createReadStream(tempDirectoryPath + '/only-headers.xlsx');
        writeWorkbooks.defaults = () => fs.createReadStream(tempDirectoryPath + '/defaults.xlsx');
        // writeWorkbooks.multiSheetNRowHeader = () => fs.createReadStream(tempDirectoryPath + '/2sheet-nheader.xlsx');

        return fs.mkdir(tempDirectoryPath);
    });

    describe('Metadata', () => {
        it('should give error if no sheet key provided for a schema', () => {
            let writer = new ExcelWriter({
                sheets: [{
                    name: 'Data',
                    key: 'data'
                }]
            });
            return writer.save()
            .then(() => {
                throw new Error('Writer must throw error');
            })
            .catch((err) => {
                expect(err).to.be.an('error');
                expect(err.message).to.match(/key is required/);
            });
        });

        it('should write headers even if no data is provided', () => {
            let writer = new ExcelWriter({
                sheets: [{
                    name: 'Data',
                    key: 'data',
                    headers: [{
                        name: 'Sr No',
                        key: 'index'
                    }, {
                        name: 'Name',
                        key: 'name'
                    }]
                }]
            });
            return writer.save()
            .then((stream) => {
                // getting the workbook stream
                let workbook = new Excel.Workbook();
                stream.pipe(workbook.xlsx.createInputStream());
                let sheet = workbook.getWorksheet('Data');
                let output = [['Sr No', 'Name']];

                expect(sheet.actualRowCount).to.equal(1);
                worksheet.eachRow((row, index) => {
                    expect(row).to.equal(row[index]);
                });
            });
        });
    });

    describe('Data', () => {
        it('should set default as empty string if no data for cell', () => {
            let writer = new ExcelWriter({
                sheets: [{
                    name: 'Data',
                    key: 'data',
                    header: [{
                        name: 'Sr No',
                        key: 'index'
                    }, {
                        name: 'Name',
                        key: 'name'
                    }, {
                        name: 'Test Value',
                        key: 'value'
                    }]
                }]
            });

            return writer.addData('data', {index: 1, value: 25})
            .then(() => {
                return writer.save();
            })
            .then((stream) => {
                let workbook = new Excel.Workbook();
                stream.pipe(workbook.xlsx.createInputStream());
                let sheet = workbook.getWorksheet('Data');
                let output = [['Sr No', 'Name', 'Test Value'], [1, '', 25]];

                expect(sheet.actualRowCount).to.equal(1);
                worksheet.eachRow((row, index) => {
                    expect(row).to.equal(output[index]);
                });
            })
        });

        it('should take default from schema if no data for cell', () => {
            let writer = new ExcelWriter({
                sheets: [{
                    name: 'Data',
                    key: 'data',
                    header: [{
                        name: 'Sr No',
                        key: 'index'
                    }, {
                        name: 'Name',
                        key: 'name',
                        default: 'Unknown'
                    }, {
                        name: 'Test Value',
                        key: 'value'
                    }]
                }]
            });

            return writer.addData('data', {index: 1, value: 25})
            .then(() => {
                return writer.save();
            })
            .then((stream) => {
                let workbook = new Excel.Workbook();
                stream.pipe(workbook.xlsx.createInputStream());
                let sheet = workbook.getWorksheet('Data');
                let output = [['Sr No', 'Name', 'Test Value'], [1, 'Unknown', 25]];

                expect(sheet.actualRowCount).to.equal(1);
                worksheet.eachRow((row, index) => {
                    expect(row).to.equal(output[index]);
                });
            })
        });

        it('should write data in some sheets', () => {
            let writer = new ExcelWriter({
                sheets: [{
                    name: 'Data',
                    key: 'data',
                    header: [{
                        name: 'Sr No',
                        key: 'index'
                    }, {
                        name: 'Name',
                        key: 'name'
                    }, {
                        name: 'Test Value',
                        key: 'value'
                    }]
                }, {
                    name: 'Second Data',
                    key: 'secondData',
                    header: [{
                        name: 'Sr No',
                        key: 'index',
                    }, {
                        name: 'Name',
                        key: 'name'
                    }]
                }]
            });

            let promises = [writer.addData('data', {index: 1, name: 'Test 1', value: 5}),
                writer.addData('data', {index: 2, name: 'Test 2', value: 15}),
            ];
            return Promise.all(promises)
            .then(() => {
                return writer.save();
            })
            .then((stream) => {
                let workbook = new Excel.Workbook();
                stream.pipe(workbook.xlsx.createInputStream());
                let sheet1 = workbook.getWorksheet('Data');
                let sheet2 = workbook.getWorksheet('Second Data');
                let firstOutput = [['Sr No', 'Name', 'Test Value'], [1, 'Test 1', 5], [2, 'Test 2', 15]];
                let secondOutput = [['Sr No', 'Name']];

                expect(sheet1.actualRowCount).to.equal(1);
                sheet1.eachRow((row, index) => {
                    expect(row).to.equal(firstOutput[index]);
                });
                sheet2.eachRow((row, index) => {
                    expect(row).to.equal(secondOutput[index]);
                });
            })
        });

        it('should write data in all sheets', () => {
        });

        it('should give error if sheet key is incorrect', () => {
        });

        it('should write the excel to a file', () => {
        });
    });

    describe('Debug', () => {
        it('should not log anything when no debug option set', () => {
        });

        it('should log when debug options is set', () => {
        });
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