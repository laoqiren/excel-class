# excel-class
a module helping to handle excel based on js-xlsx

### getting start

`
npm install excel-class
`
### APIs

#### get the excel object and start to use it

```
const Excel = require('excel-class');
const path = require('path');
let excel = new Excel(path.join(__dirname,'test.xlsx')
```
**tip: the path of the excel file should be absolute path**

#### Read sheet by sheetName