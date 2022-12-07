var Excel = require('exceljs')
var XLSX = require('xlsx');
function getHeadersFromExcel(){
    var wb = XLSX.readFile('testdata.xlsx');
        var ws = wb.Sheets.Sheet1;
        var xlData =XLSX.utils.sheet_to_json(wb.Sheets.Sheet1, {header:1})
        //console.log(xlData);
        var header_Data=xlData[0]
        return header_Data
    }

var testdata=[];
 
var header_values=getHeadersFromExcel()
 

    var resultset=generateRandamData(header_values);
    writeDataToExcel(resultset)
    console.log("written to excel")


    

 function writeDataToExcel(resultset)
 {
    
           var workbook = new Excel.Workbook()
           //var arr=[]
           workbook.xlsx.readFile('./testdata.xlsx')
           .then(function(){
            var worksheet = workbook.getWorksheet(1)
          // Add an array of rows

    
     
  // add new rows and return them as array of row objects
  const newRows = worksheet.addRows(resultset);
 
  
   workbook.xlsx.writeFile('./testdata.xlsx')
           })
 }   




function randnumber(length) {
    var result           = '';
    var characters       = '1234567890';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
        result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }
    return result;
}
 
function randastring(length) {
    var result           = '';
    var characters       = 'abcdefghijklmnopqrstuvwxyz';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
        result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }
    return result;
}
 
function randaalphanumeric(length) {
    var result           = '';
    var characters       = 'abcdefghijklmnopqrstuvwxyz123456789';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
        result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }
    return result;
}
function generateRandamData(xldata)
{
    var arr=new Array()
    for(let i=0;i<xldata.length;i++)
    {
        if(xldata[i]=='name')
        {
            testdata.push(randastring(8));
 
        }
        else if(xldata[i]=='age')
        {
            testdata.push(randnumber(2));
            
        }
        else if(xldata[i]=='phone')
        {
            testdata.push(randnumber(10));
           
        }
        else if(xldata[i]=='address')
        {
            testdata.push(randaalphanumeric(15));
           
        }
        else if(xldata[i]=='contactid')
        {
            testdata.push(randnumber(6));
           
        }
    }
    arr.push(testdata)
    testdata=[]
    return arr
}
