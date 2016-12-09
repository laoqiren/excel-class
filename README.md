# excel-class
a module helping to handle excel based on js-xlsx

### Getting start

`
npm install excel-class
`
### APIs

#### Get the excel object and start to use it

```
const Excel = require('excel-class');
const path = require('path');
let excel = new Excel(path.join(__dirname,'test.xlsx')
```
**tip: the path of the excel file should be absolute path**

#### readSheet(sheet)
the argument can be type of Number or String

Read sheet by sheetName or sheetNumber, it will return json includes all data in the sheet.

```
excel.readSheet('Sheet1');
excel.readSheet(0);
```

#### readRow(sheet,rowNumber)
read the excel by rowNumber, the rowNumber should must largger than 0, if rowNumber is 0, this api will return a array of headers.

#### readCell(sheet,rowNumber,cell)

return the correct cell of data, the cell can be type of Number or String.

```
excel.readCell('Sheet1',1,5);
excel.readCell('Sheet1',1,'name')
```

#### writeSheet(sheet,headers,data)

write a new sheet to the excel, if the sheet has exsited, the new data will replace it. And the other sheets in the excel will not changed.

headers shoud be an array contains the headers of this sheet, and data should be type of array and it contains data of rows:

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
]);
```

#### writeRow(sheet,row,data)

write the data of row in the sheet, the data should be a object

```
excel.writeRow('Sheet1',1,{
    name: 'Jane',
    age: 19,
    country: 'China'
})
```