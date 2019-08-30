let path = require("path");
let { excelRead } = require("./index.js");

excelRead(path.join(".","example.xlsx"),1,2,5).then(datas=>{
	console.log(datas);
});




