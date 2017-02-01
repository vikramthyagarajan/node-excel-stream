# node-excel-stream
A utility to read and write excels using streams in Node

This package provides ExcelWriter and ExcelReader classes that can be used to writer and parse excel respectively

## Installation
```
npm install node-excel-stream
```
Usage with es5
```
// with es5
var ExcelReader = require('node-excel-stream').ExcelReader;
var ExcelWriter = require('node-excel-stream').ExcelWriter;
// or with es6 imports
import { ExcelReader, ExcelWriter } from 'node-excel-stream';
```

## Usage

### Parsing Workbooks
Workbooks must be given a schema before they are parsed. If the excel format does not match the given schema, then an error is thrown. Then, each row in each sheet can be parsed using the eachRow function, which returns a promise when resolves after all rows are parsed.

Input: data.xlsx, Sheet: Users

| User Name | Value |
|:----------|:-----:|
| John | 10 |
| Rohan | 30 |
| Pooja | 50 |

Parsing:
```
let dataStream = fs.createReadStream('data.xlsx');
let reader = new ExcelReader(dataStream, {
    sheets: [{
        name: 'Users',
        allowedNames: ['Users'],
        rows: {
            headerRow: 1,
            allowedHeaders: [{
                name: 'User Name',
                key: 'userName'
            }, {
                name: 'Value',
                key: 'value',
                type: Number
            }]
        }
    }]
})
console.log('starting parse');
reader.eachRow((rowData, rowNum, sheetSchema) => {
    console.log(rowData);
})
.then(() => {
    console.log('done parsing');
});
```

Output:
```
starting parse
{ userName: 'John', value: 10 }
{ userName: 'Rohan', value: 10 }
{ userName: 'Pooja', value: 10 }
finished parse
```

### Writing Workbooks
Workbooks must be given a schema for each sheet before it can be written to. Data is then added to the workbook in the form of json. All this data is saved to the sheet using the save function, which returns a promise which resolves the stream. This stream can then be piped or written to the fileSystem using the fs module.

Input: var inputs
```
[
    {name: 'Test 1', testValue: 100},
    {name: 'Test 2'},
    {name: 'Test 3', testValue: 80}
]
```

Writing:
```
let writer = new ExcelWriter({
    sheets: [{
        name: 'Test Sheet',
        key: 'tests',
        headers: [{
            name: 'Test Name',
            key: 'name'
        }, {
            name: 'Test Coverage',
            key: 'testValue',
            default: 0
        }]
    }]
});
let dataPromises = inputs.map((input) => {
    // 'tests' is the key of the sheet. That is used
    // to add data to only the Test Sheet
    writer.addData('tests', input);
});
Promise.all(dataPromises)
.then(() => {
    return writer.save();
})
.then((stream) => {
    stream.pipe(fs.createWriteStream('data.xlsx'));
});
```

Output:
Input: data.xlsx, Sheet: Test Sheet

| Test Name | Test Coverage |
|:----------|:-----:|
| Test 1 | 100 |
| Test 2 | 0 |
| Test 3 | 80 |
