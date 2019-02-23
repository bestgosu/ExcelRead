let path = require("path");
let excelRead = require("./ExcelRead.js");

excelRead(path.join(".","기업은행.xls"),2,3,50).then(datas=>{
	console.log(datas);
});




