# excel-class

[![npm](https://img.shields.io/npm/dm/excel-class.svg?style=flat-square)](https://www.npmjs.com/package/excel-class)
[![npm](https://img.shields.io/npm/v/excel-class.svg?style=flat-square)](https://github.com/laoqiren/excel-class)
[![Travis](https://img.shields.io/travis/rust-lang/rust.svg?style=flat-square)]()

a module helping to handle excel based on js-xlsx

[中文文档](https://github.com/laoqiren/excel-class/blob/master/CN.md)
### Getting start

`
npm install excel-class
`
### APIs

### Get the excel object and start to use it

```
const Excel = require('excel-class');
const path = require('path');
let excel = new Excel(path.join(__dirname,'test.xlsx')
```
**tip: the path of the excel file should be absolute path**

### readSheet(sheet)
the argument can be type of Number or String

Read sheet by sheetName or sheetNumber, it will return json includes all data in the sheet.

```
excel.readSheet('Sheet1');
excel.readSheet(0);
```

### readRow(sheet,rowNumber)
read the excel by rowNumber, the rowNumber should must largger than 0, if rowNumber is 0, this api will return an array contains headers.

### readCell(sheet,rowNumber,cell)

return the correct cell of data, the cell can be type of Number or String.

```
excel.readCell('Sheet1',1,5);
excel.readCell('Sheet1',1,'name')
```

### writeSheet(sheet,headers,data)

write a new sheet to the excel, if the sheet has exsited, the new data will replace it. And the other sheets in the excel will not changed.

headers shoud be an array contains the headers of this sheet, and data should be type of array and it contains data of rows. This api will return a promise.

```
excel.writeSheet('Sheet1',['name','age','country'],[
    {
        name: 'Jane',
        age: 19,
        country: 'China'
    },
    {
        name: 'Maria',
        age: 20,
        country: 'America'
    }
]).then(()=>{
    //do other things
});
```

### writeRow(sheet,row,data)

write the data of row in the sheet, the data should be a object, this api will return a promise

```
excel.writeRow('Sheet1',1,{
    name: 'Jane',
    age: 19,
    country: 'China'
}).then(()=>{
    //do other things
})
```
