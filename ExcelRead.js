let path = require("path");
//let fs = require("fs");

let _r =require("ramda");
let XLSX = require("xlsx");

let F = function(f){
	return f;
}

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

async function excelReadByBinaryString(bstr,nameRowNum,dataStartNum,dataEndNum){
	let workbook = XLSX.read(bstr,{type:'binary'});
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

function getLastRowNum(worksheet){
	let lastKey = _r.last(getDataKeys(worksheet));
	//let lastKey = "A101010"
	//getNumberString from lastKey
	return Number(_r.match(new RegExp("[0-9]+","ig"),lastKey)[0]);
}

function getExcelArray(filePath){
	let workbook = await XLSX.readFile(filePath);
	let worksheet_name = workbook.SheetNames[0];
	let worksheet = workbook.Sheets[worksheet_name];

	let lastNum = getLastRowNum(worksheet);
	let rowRange = _r.range(1,lastNum);
	let excelArray = _r.map(num=>getRowDatas(worksheet,num))(rowRange);
	return excelArray;
}

module.exports = { excelRead, excelReadByBinaryString,getExcelArray } ;


//if main execute this
if (require.main === module) {
	//execute functions

	async function main() {
		let workbook = await XLSX.readFile(path.join(".", "write.xlsx"));
		let worksheet_name = workbook.SheetNames[0];
		let worksheet = workbook.Sheets[worksheet_name];

		//function name and function execute console
		console.log("worksheet",worksheet)
		console.log("getDataKeys", getDataKeys(worksheet))
		console.log("getRowKeys(1)", getRowKeys(worksheet, 1))
		console.log("getColumnKeys(A)", getColumnKeys(worksheet, "A"))
		console.log("getRowDatas(1)", getRowDatas(worksheet, 1))
		console.log("getColumnDatas(A)", getColumnDatas(worksheet, "A"))

		console.log(await excelRead(path.join(".", "write.xlsx"),1,2,5))

		console.log(getExcelArray(worksheet))

	}
	main();
}
