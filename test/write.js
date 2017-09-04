const Excel = require('../index');
const path = require('path');
const expect = require('chai').expect;

describe("excel.writeSheet()",function(){
    it("write a new sheet to the excel while the file hasn't exsited",function(done){
        let excel = new Excel(path.join(__dirname,'../assets/new.xlsx'));

        excel.writeSheet("Sheet1",['name','age','country'],[
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
            done();
        }).catch(err=>{
            done(err);
        })
    })

    it("replace the exsited sheet",function(done){
        let excel = new Excel(path.join(__dirname,'../assets/test1.xlsx'));

        excel.writeSheet("Sheet1",['name','age','country'],[
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
            done();
        }).catch(err=>{
            done(err);
        })
    })
});

describe("excel.writeRow()",function(){
    it("write data to specify row",function(done){
        let excel = new Excel(path.join(__dirname,'../assets/test1.xlsx'));

        excel.writeRow('Sheet1',1,{
            name: 'Jane',
            age: 19,
            country: 'China'
        }).then(()=>{
            done();
        }).catch(err=>{
            done(err);
        })
    })
})

describe("excel.writeCell()",function(){
    it("write data to specify column",function(done){
        let excel = new Excel(path.join(__dirname,'../assets/test1.xlsx'));

        excel.writeCell('Sheet1',1,'age',20).then(()=>{
            done();
        }).catch(err=>{
            done(err);
        })
    })
})