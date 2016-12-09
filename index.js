const XLSX = require('xlsx');
const fs = require('fs');
class Excel{
    constructor(file){
        this.excel = file;
    }
    readFile(){
        return XLSX.readFile(this.excel);
    }
    readSheet(sheet){
        let excel = this.readFile();
        if(typeof sheet === 'number'){
            return XLSX.utils.sheet_to_json(excel.Sheets[excel.SheetNames[sheet]]);
        } else if(typeof sheet === 'string'){
            return XLSX.utils.sheet_to_json(excel.Sheets[sheet]);
        } else {
            return null;
        }
    }
    readHead(sheet){
        let workSheet = XLSX.readFile(this.excel,{
            sheetRows:1
        }).Sheets[sheet];
        let result = [];
        for(let z in workSheet){
            if(z[0] === '!') continue;
            result.push(workSheet[z].v);
        }
        return result;
    }
    readRow(sheet,row){
        let rows = this.readSheet(sheet);
        if(row === 0){
            return this.readHead(sheet);
        }
        return rows[row-1];
    }
    readCell(sheet,rowNumber,column){
        let row = this.readRow(sheet,rowNumber);
        if(typeof column === 'number'){
            return row[this.readHead(sheet)[column-1]]
        } else if(typeof column === 'string'){
            return row[column]
        } else {
            return undefined;
        }
    }
    writeSheet(sheet,headers,data){
        let _headers = headers.map((value,index)=>
            Object.assign({},{
                v:value,
                position: String.fromCharCode(65+index) + 1
            })
        ).reduce((pre,next)=>Object.assign({},pre,{[next.position]:{v:next.v}}),{})
        let _data = data
            .map((v,i)=>{
                return headers.map((hv,hi)=>{
                    return Object.assign({},{
                        v:v[hv],
                        position: String.fromCharCode(65+hi) + (i+2)
                    })
                })
            })
            .reduce((pre,next)=>pre.concat(next))
            .reduce((pre,next)=>Object.assign({},pre,{
                [next.position]:{v:next.v}
            }),{})
        let output = Object.assign({},_headers,_data),
            pos = Object.keys(output),
            ref = pos[0] + ':' + pos[pos.length-1];

        this.writeExcel(output,sheet,ref);
        
    }
    writeExcel(newSheet,sheetName,ref){
        let excel = this.readFile();
        let workbook;
        fs.stat(this.excel,(err,stat)=>{
            if(stat && stat.isFile()){
                excel.SheetNames.push(sheetName);
                workbook = {
                    SheetNames: [...new Set(excel.SheetNames)],
                    Sheets: Object.assign(excel.Sheets,{
                        [sheetName]: Object.assign({},{'!ref':ref},newSheet)
                    })
                }
            } else {
                workbook = {
                    SheetNames: [sheetName],
                    Sheets: {
                        [sheetName]: Object.assign({},{'!ref':ref},newSheet)
                    }
                }
            }
            try {
                XLSX.writeFile(workbook,this.excel)
            } catch(err){
                if(err){
                    console.error(err)
                    return false;
                }
            }
            console.log('写入excel表格成功！')
        });
    }
    writeRow(sheet,row,data){
        let preData = this.readSheet(sheet),
            headers = this.readHead(sheet);
        preData[row-1] = data;
        this.writeSheet(sheet,headers,preData);
    }
    writeCell(sheet,row,column,data){
        let preData = this.readSheet(sheet),
            headers = this.readHead(sheet);
        preData[row-1][column] = data;
        this.writeSheet(sheet,headers,preData);
    }
}

module.exports = Excel;