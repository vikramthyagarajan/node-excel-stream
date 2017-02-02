'use strict';
const expect = require('chai').expect;
const Excel = require('exceljs');
const ExcelReader = require('../index').ExcelReader;
const fs = require('fs');

let testWorkbooks = {};
describe('Excel Reader', () => {
    before(() => {
        testWorkbooks.singleSheetFirstRowHeader = () => fs.createReadStream(__dirname + '/util/excels/1sheet-1header.xlsx');
        testWorkbooks.singleSheetNRowHeader = () => fs.createReadStream(__dirname + '/util/excels/1sheet-nheader.xlsx');
        testWorkbooks.multiSheetNRowHeader = () => fs.createReadStream(__dirname + '/util/excels/2sheet-nheader.xlsx');
    });
    describe('Sheets', () => {
        it('should error if different number of Sheets', () => {
            let reader = new ExcelReader(testWorkbooks.singleSheetFirstRowHeader(), {
                sheets: [{
                    name: 'Data',
                }, {
                    name: 'Test'
                }]
            })
            return reader.eachRow()
            .then(() => {
                throw 'Reader must exit with an error';
            })
            .catch((err) => {
                expect(err).to.be.an('error');
                expect(err.message).to.match(/Schema not found/i);
            });
        });
    });

    describe('Allowed Sheet Names', () => {
        it('should be an array or null', () => {
            let reader = new ExcelReader(testWorkbooks.singleSheetFirstRowHeader(), {
                sheets: [{
                    allowedNames: 'dkfjkdj'
                }]
            });
            return reader.eachRow()
            .then(() => {
                done('Validation should give an error');
            })
            .catch((err) => {
                expect(err).to.be.an('error');
                expect(err.message).to.match(/"allowedNames" must be an array/);
            })
        });

        it('should only allow selected sheet names', () => {
            let reader = new ExcelReader(testWorkbooks.singleSheetFirstRowHeader(), {
                sheets: [{
                    name: 'Data'
                }]
            });
            return reader.eachRow()
            .then(() => {
                throw new Error('Reader must throw error');
            })
            .catch((err) => {
                expect(err).to.be.an('error');
                expect(err.message).to.match(/Schema not found/);
            });
        });

        it('should throw error if invalid sheet name is in excel', () => {
            let reader = new ExcelReader(testWorkbooks.singleSheetFirstRowHeader(), {
                sheets: [{
                    name: 'Incorrect Sheet'
                }]
            });
            return reader.eachRow()
            .then(() => {
                throw new Error('Reader must throw error');
            })
            .catch((err) => {
                expect(err).to.be.an('error');
                expect(err.message).to.match(/Schema not found/);
            });
        });
    });

    describe('Rows', () => {
        it('should assume default header row', (done) => {
            let reader = new ExcelReader(testWorkbooks.singleSheetFirstRowHeader, {
                sheets: [{
                    name: 'Data',
                    rows: {
                        allowedHeaders: [{
                            name: 'Sr No',
                            key: 'index'
                        }, {
                            name: 'Name',
                            key: 'name'
                        }, {
                            name: 'X Value',
                            key: 'x'
                        }, {
                            name: 'Y Value',
                            key: 'y'
                        }, {
                            name: 'Z Value',
                            key: 'z'
                        }, {
                            name: 'Total',
                            key: 'total'
                        }]
                    }
                }]
            });
            let output = [{
                index: '1',
                name: 'First Entry',
                x: '25',
                y: '30',
                z: '45',
                total: '100'
                }, {
                    index: '2',
                    name: 'Second Entry',
                    x: '20',
                    y: '20',
                    z: '20',
                    total: '60'
                }, {
                    index: '3',
                    name: 'Third Entry',
                    x: '15',
                    y: '10',
                    z: '8',
                    total: '33'
                }, {
                    index: '4',
                    name: 'Fourth Entry',
                    x: '22',
                    y: '39',
                    z: '65',
                    total: '126'
                }, {
                    index: '5',
                    name: 'Fifth Entry',
                    x: '42',
                    y: '8',
                    z: 'invalid num',
                    total: '50'
                }];

            return reader.eachRow((rowData, rowNum) => {
                expect(rowData).to.eql(output[rowNum - 1]);
            });
        });

        it('should take header row from config', (done) => {
            let reader = new ExcelReader(testWorkbooks.singleSheetNRowHeader, {
                sheets: [{
                    name: 'Data',
                    rows: {
                        headerRow: 6,
                        allowedHeaders: [{
                            name: 'Sr No',
                            key: 'index'
                        }, {
                            name: 'Name',
                            key: 'name'
                        }, {
                            name: 'X Value',
                            key: 'x'
                        }, {
                            name: 'Y Value',
                            key: 'y'
                        }]
                    }
                }]
            });
            let output = [{
                    index: '1',
                    name: 'First Entry',
                    x: '6',
                    y: '68'
                }, {
                    index: '2',
                    name: 'Second Entry',
                    x: '34',
                    y: '57'
            }];

            return reader.eachRow((rowData, rowNum) => {
                expect(rowData).to.eql(output[rowNum - 1]);
            });
        });
        
        it('should return each row data for multiple sheets', (done) => {
            let reader = new ExcelReader(testWorkbooks.multiSheetNRowHeader, {
                sheets: [{
                    name: 'First Sheet',
                    key: 'sheet1',
                    rows: {
                        headerRow: 4,
                        allowedHeaders: [{
                            name: 'Sr No',
                            key: 'index'
                        }, {
                            name: 'Name',
                            key: 'name'
                        }, {
                            name: 'X Value',
                            key: 'x'
                        }, {
                            name: 'Y Value',
                            key: 'y'
                        }, {
                            name: 'Total',
                            key: 'total'
                        }]
                    }
                }, {
                    name: 'Second Sheet',
                    key: 'sheet2',
                    rows: {
                        headerRow: 3,
                        allowedHeaders: [{
                            name: 'Name',
                            key: 'name'
                        }, {
                            name: 'Z Value',
                            key: 'x'
                        }, {
                            name: 'Total',
                            key: 'total'
                        }]
                    }
                }]
            });
            let output = {
                sheet1: [{
                        index: '1',
                        name: 'First Entry',
                        x: '25',
                        y: '5',
                        total: '30'
                    }, {
                        index: '2',
                        name: 'Second Entry',
                        x: '20',
                        y: '20',
                        total: '40'
                    }, {
                        index: '3',
                        name: 'Third Entry',
                        x: '15',
                        y: '10',
                        total: '25'
                    }, {
                        index: '4',
                        name: 'Fourth Entry',
                        x: '22',
                        y: 'error',
                        total: '22'
                }],
                sheet2: [{
                        name: 'First Entry',
                        z: '43',
                        total: '73'
                    }, {
                        name: 'Second Entry',
                        z: '77',
                        total: '117'
                    }, {
                        name: 'Second Entry',
                        z: '51',
                        total: '76'
                }]
            };;

            reader.eachRow((rowData, rowNum, sheetKey) => {
                expect(rowData).to.eql(output[sheetKey][rowNum - 1]);
            });
        });
    });
});