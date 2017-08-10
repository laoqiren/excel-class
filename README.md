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
let excel = new Excel(path.join(__dirname,'test.xlsx'))
```
**tip: the path of the excel file should be absolute path**

### readSheet(sheet)
get the data of specify sheet

Read sheet by sheetName or sheetNumber, it will return json includes all data in the sheet.

* **sheet** \<Number\>|\<String\> the number of sheet or the name of sheet to read.
* Returns: \<Array\> all the data in the sheet.


```
excel.readSheet('Sheet1');
excel.readSheet(0);
```

### readRow(sheet,row)

get the specify row's value, row should largger than 0, if row is 0, this api will return an array contains headers.

* **row** \<Number\> the number of row to read.
* Returns: \<String\> the result

### readCell(sheet,row,column)

get the specify column's value, column can be type of Number or String.

* **column** \<Number\>|\<String\> the number or name of column to read.
* Returns: \<String\> the result 

```
excel.readCell('Sheet1',1,5);
excel.readCell('Sheet1',1,'name')
```

### writeSheet(sheet,headers,data)

write a new sheet to the excel, if the sheet has exsited, the new data will replace it. And the other sheets in the excel will not changed.

headers shoud be an array contains the headers of this sheet, and data should be type of array and it contains data of rows. This api will return a promise.

* **headers** \<Array\> the headers of the sheet
* **data** \<Array\> the data to write
* Returns: Promise Object

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

write data to specify row , the data should be a object, this api will return a promise

* **row** \<Number\> the number of row to write.
* **data** \<Object\> the data to write
* Returns: Promise Object

```
excel.writeRow('Sheet1',1,{
    name: 'Jane',
    age: 19,
    country: 'China'
}).then(()=>{
    //do other things
})
```

### writeCell(sheet,row,column,data)

write data to specify column , the data should be a object, this api will return a promise

* **column** \<String\> the name of cloumn to write.
* **data** the data to write
* Returns: Promise Object

```
excel.writeCell('Sheet1',1,'age',20).then(()=>{
    //do other things
})
```

### LICENSE
MIT.
