let path = require("path");
//let fs = require("fs");

let _r =require("ramda");
let XLSX = require("xlsx");



//make function json to excel
function excelWrite(json,fileName){
  let ws = XLSX.utils.json_to_sheet(json);
  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Sheet1");
  XLSX.writeFile(wb,fileName);
}


//export
module.exports = { excelWrite };

//if main execute this
if (require.main === module) {
    let json = [
        {"이름":"홍길동","나이":20,"성별":"남"},
        {"이름":"김길동","나이":30,"성별":"여"},
        {"이름":"박길동","나이":40,"성별":"남"}
    ];

    excelWrite(json, path.join(".", "write.xlsx"));
}