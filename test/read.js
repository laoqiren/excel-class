const Excel = require('../index');
const path = require('path');
const expect = require('chai').expect;

describe("excel",function(){
    it("create operation object without error",function(){
        let excel = new Excel(path.join(__dirname,'../assets/test.xlsx'));
        expect(excel).to.be.an('object');
    })
});

describe("excel.readSheet()",function(){
    it("read by Sheet name",function(){
        let excel = new Excel(path.join(__dirname,'../assets/test.xlsx'));

        let result = excel.readSheet("Sheet1");
        expect(result).to.be.an('array');
    })

    it("read by Sheet Number",function(){
        let excel = new Excel(path.join(__dirname,'../assets/test.xlsx'));

        let result = excel.readSheet(0);
        expect(result).to.be.an('array');
    })
})

describe("excel.readRow()",function(){
    it("read specify row",function(){
        let excel = new Excel(path.join(__dirname,'../assets/test.xlsx'));

        let result = excel.readRow("Sheet1",2);
        expect(result).to.be.an('object');
    })
})

describe("excel.readCell()",function(){
    it("read specify cell",function(){
        let excel = new Excel(path.join(__dirname,'../assets/test.xlsx'));

        let result = excel.readCell("Sheet1",2,3);
        expect(result).to.be.a('string');
    })
})