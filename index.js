let XLSX = require("xlsx");
let workbook = XLSX.readFile("sample.xlsx");

let sheet_name_list = workbook.SheetNames;
let Sheet1 = workbook.Sheets[sheet_name_list[0]];
let Sheet1_json = XLSX.utils.sheet_to_json(Sheet1);

let Sheet1A1 = Sheet1["E4"].v;
console.log( `シート1のセルE4の値：\n${Sheet1A1}`);

console.log("sheet_name_list:", sheet_name_list);
console.log(Sheet1_json);

function handleFile(e) {
    let files = e.target.files, f = files[0];
    let reader = new FileReader();
    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, {type: "array"});
    };
    reader.readAsArrayBuffer(f);
}

input_dom_element.addE