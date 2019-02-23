let path = require("path");
let fs = require("fs");

//let { F } = require("../safeWrapper");
let F = function(f){
	return f;
};

let _r =require("ramda");
let XLSX = require("xlsx");

async function excelRead(filePath,nameRowNum,dataStartNum,dataEndNum){
	let workbook = await XLSX.readFile(filePath);
	let worksheet_name = workbook.SheetNames[0];
	let worksheet = workbook.Sheets[worksheet_name];

	let nameRow = _r.compose(
		_r.map(key=>getValue(worksheet,key))
		//,_r.tap(obj=>console.log(JSON.stringify(obj)))
		,num=>F(getRowKeys)(worksheet,num)
	)(nameRowNum);

	let dataRowRange = _r.range(dataStartNum,dataEndNum);
	let data = _r.map(num=>{
		return _r.compose(
			_r.zipObj(nameRow)
			,_r.map(key=>getValue(worksheet,key))
			//,_r.tap(obj=>console.log(JSON.stringify(obj)))
			,num=>F(getRowKeys)(worksheet,num)
		)(num);
	})(dataRowRange);

	return data;
}
function getDataKeys(worksheet){
	let keys = _r.keys(worksheet);
	let dataKeys = _r.filter(key=>{
		return _r.test(new RegExp("^[a-zA-z]+[0-9]+","i"))(key);
	},keys);

	return dataKeys;
}
function getValue(worksheet,cell){
	return worksheet[cell].v;
}
function getRowKeys(worksheet,rowNum){
	return _r.compose(
		_r.filter((key)=>{
			return _r.test(new RegExp(`^[a-zA-Z]+${rowNum}$`,"ig"))(key);
		})
		,getDataKeys
	)(worksheet);
}
function getColumnKeys(worksheet,columnChar){
	return _r.compose(
		_r.filter((key)=>{
			return _r.test(new RegExp(`${columnChar}[0-9]+`,"ig"))(key);
		})
		,getDataKeys
	)(worksheet);
}
function getRowDatas(worksheet,rowNum){
	return _r.map(key=>getValue(worksheet,key))(getRowKeys(worksheet,rowNum))
}
function getColumnDatas(worksheet,columnChar){
	return _r.map(key=>getValue(worksheet,key))(getColumnKeys(worksheet,columnChar))
}

module.exports = excelRead;
