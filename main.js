let _r = require("ramda");

let path = require("path");
let { excelRead, excelWrite } = require("./index.js");

excelRead(path.join(".", "write.xlsx"), 1, 2, 5).then(
	datas => {
		_r.filter(data => data["이름"] === "김길동")(datas).forEach(data => {
			console.log(data);
		})
});





