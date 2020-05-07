'use strict';
const expect = require('chai').expect;
const Excel = require('exceljs');
const ExcelWriter = require('../index').ExcelWriter;
const fs = require('fs');
const rm = require('rimraf');

let readWorkbooks = {}, writeWorkbooks = {};
let tempDirectoryPath = __dirname + '/temp';
function cleanValues(arr) {
    // row.values gives the first cell as null, so stripping that
    return arr.slice(1);
}
describe('Excel Writer', () => {
    before(() => {
        // before the tests here, we must create a temp directory, and write the excels there
        // When checking the excels is required, then a read stream is returned for that file
        // using these functions-
        readWorkbooks.multiSheet = () => fs.createReadStream(tempDirectoryPath + '/multi-sheet.xlsx');

        writeWorkbooks.multiSheet = () => fs.createWriteStream(tempDirectoryPath + '/multi-sheet.xlsx');

        return fs.mkdir(tempDirectoryPath, () => {});
    });

    describe('Metadata', () => {
        it('should give error if no sheet key provided for a schema', () => {
            let writer = new ExcelWriter({
                sheets: [{
                    name: 'Data'
                }]
            });
            return writer.save()
            .then(() => {
                throw new Error('Writer must throw error');
            })
            .catch((err) => {
                expect(err).to.be.an('error');
                expect(err.message).to.match(/no key specified/i);
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
                return workbook.xlsx.read(stream)
                .then((workbook) => {
                    let sheet = workbook.getWorksheet('Data');
                    let output = [['Sr No', 'Name']];

                    expect(sheet.actualRowCount).to.equal(1);
                    sheet.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(output[index - 1]);
                    });
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
                    headers: [{
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
                return workbook.xlsx.read(stream)
                .then((workbook) => {
                    let sheet = workbook.getWorksheet('Data');
                    let output = [['Sr No', 'Name', 'Test Value'], [1, , 25]];

                    expect(sheet.actualRowCount).to.equal(output.length);
                    sheet.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(output[index - 1]);
                    });
                });
            })
        });

        it('should take default from schema if no data for cell', () => {
            let writer = new ExcelWriter({
                sheets: [{
                    name: 'Data',
                    key: 'data',
                    headers: [{
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
                return workbook.xlsx.read(stream)
                .then((workbook) => {
                    let sheet = workbook.getWorksheet('Data');
                    let output = [['Sr No', 'Name', 'Test Value'], [1, 'Unknown', 25]];

                    expect(sheet.actualRowCount).to.equal(output.length);
                    sheet.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(output[index - 1]);
                    });
                });
            })
        });

        it('should write data in some sheets', () => {
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
                    }, {
                        name: 'Test Value',
                        key: 'value'
                    }]
                }, {
                    name: 'Second Data',
                    key: 'secondData',
                    headers: [{
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
                return workbook.xlsx.read(stream)
                .then((workbook) => {
                    let sheet1 = workbook.getWorksheet('Data');
                    let sheet2 = workbook.getWorksheet('Second Data');
                    let firstOutput = [['Sr No', 'Name', 'Test Value'], [1, 'Test 1', 5], [2, 'Test 2', 15]];
                    let secondOutput = [['Sr No', 'Name']];

                    expect(sheet1.actualRowCount).to.equal(firstOutput.length);
                    expect(sheet2.actualRowCount).to.equal(secondOutput.length);
                    sheet1.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(firstOutput[index - 1]);
                    });
                    sheet2.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(secondOutput[index - 1]);
                    });
                });
            })
        });

        it('should write data in all sheets', () => {
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
                    }, {
                        name: 'Test Value',
                        key: 'value'
                    }]
                }, {
                    name: 'Second Data',
                    key: 'secondData',
                    headers: [{
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
                writer.addData('secondData', {index: 1, name: 'Test 21'}),
            ];
            return Promise.all(promises)
            .then(() => {
                return writer.save();
            })
            .then((stream) => {
                let workbook = new Excel.Workbook();
                return workbook.xlsx.read(stream)
                .then((workbook) => {
                    let sheet1 = workbook.getWorksheet('Data');
                    let sheet2 = workbook.getWorksheet('Second Data');
                    let firstOutput = [['Sr No', 'Name', 'Test Value'], [1, 'Test 1', 5], [2, 'Test 2', 15]];
                    let secondOutput = [['Sr No', 'Name'], [1, 'Test 21']];

                    expect(sheet1.actualRowCount).to.equal(firstOutput.length);
                    expect(sheet2.actualRowCount).to.equal(secondOutput.length);
                    sheet1.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(firstOutput[index - 1]);
                    });
                    sheet2.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(secondOutput[index - 1]);
                    });
                });
            })
        });

        it('should give error if sheet key is incorrect', () => {
            let writer = new ExcelWriter({
                sheets: [{
                    name: 'Data',
                    key: 'data'
                }]
            });
            return writer.addData('notExists', {})
            .then(() => {
                throw new Error('Writer must throw error');
            })
            .catch((err) => {
                expect(err).to.be.an('error');
                expect(err.message).to.match(/no such sheet key/i);
            });
        });

        it('should write the excel to a file', () => {
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
                    }, {
                        name: 'Test Value',
                        key: 'value'
                    }]
                }, {
                    name: 'Second Data',
                    key: 'secondData',
                    headers: [{
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
                writer.addData('secondData', {index: 1, name: 'Test 21'}),
            ];
            return Promise.all(promises)
            .then(() => {
                return writer.save();
            })
            .then((stream) => {
                let writeStream = writeWorkbooks.multiSheet();
                stream.pipe(writeStream);
                return new Promise((resolve, reject) => {
                    writeStream.on('finish', resolve);
                    writeStream.on('error', reject);
                });
            })
            .then(() => {
                let workbook = new Excel.Workbook();
                let stream = readWorkbooks.multiSheet();
                return workbook.xlsx.read(stream)
                .then((workbook) => {
                    let sheet1 = workbook.getWorksheet('Data');
                    let sheet2 = workbook.getWorksheet('Second Data');
                    let firstOutput = [['Sr No', 'Name', 'Test Value'], [1, 'Test 1', 5], [2, 'Test 2', 15]];
                    let secondOutput = [['Sr No', 'Name'], [1, 'Test 21']];

                    expect(sheet1.actualRowCount).to.equal(firstOutput.length);
                    expect(sheet2.actualRowCount).to.equal(secondOutput.length);
                    sheet1.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(firstOutput[index - 1]);
                    });
                    sheet2.eachRow((row, index) => {
                        expect(cleanValues(row.values)).to.eql(secondOutput[index - 1]);
                    });
                });
            })
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
