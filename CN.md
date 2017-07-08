# excel-class
[![npm](https://img.shields.io/npm/dm/excel-class.svg?style=flat-square)](https://www.npmjs.com/package/excel-class)
[![npm](https://img.shields.io/npm/v/excel-class.svg?style=flat-square)](https://github.com/laoqiren/excel-class)
[![Travis](https://img.shields.io/travis/rust-lang/rust.svg?style=flat-square)]()


封装excel文件常用操作，基于js-xlsx模块

[English doc](https://github.com/laoqiren/excel-class/blob/master/README.md)
### 起步
```
npm install excel-class
```

APIs

### 依赖模块，并实例化一个excel操作对象

```
const Excel = require('excel-class');
const path = require('path');
let excel = new Excel(path.join(__dirname,'test.xlsx'))
```
**注意** 文件路径名必须是绝对路径

### readSheet(sheet)

读取指定sheet,sheet参数可以是数字或者字符串，返回excel文件指定sheet的所有数据的json格式

```
excel.readSheet('Sheet1');
excel.readSheet(0);
```

### readRow(sheet,rowNumber)

读取指定sheet中指定行的数据，rowNumber取值应该大于等于0，如果等于0，会返回该sheet的headers数组

### readCell(sheet,rowNumber,cell)

返回指定行列的数据，cell可以是字符串，也可以是数字
```
excel.readCell('Sheet1',1,5);
excel.readCell('Sheet1',1,'name')
```
### writeSheet(sheet,headers,data)

新建或者替换excel中的某sheet的数据，headers是表头数组，data是对象数组，数组中每个对象表示某一行的数据，API会返回promise对象
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

写入指定行数据,API会返回promise对象

```
excel.writeRow('Sheet1',1,{
    name: 'Jane',
    age: 19,
    country: 'China'
}).then(()=>{
    //do other things
})
```

### LICENSE
MIT.
