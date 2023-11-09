const xlsx =require("xlsx"); 
const path = require('path');
let _source_file = "./data.xlsx";//Driving-Test-Canada.xls";
let _source_file_path = path.resolve(_source_file);

const workbook = xlsx.readFile(_source_file_path)
let worksheet = workbook.Sheets['AB'];

/*
let cell = worksheet['D3'];//worksheet['D3'].v; // Column C, Row 3 (assuming there's data there)

let _pre_Quote = cell.f.toString().indexOf('"');
let _post_Quote = cell.f.toString().indexOf('"', _pre_Quote+1);

let _img = cell.f.toString().substring(_pre_Quote+1, _post_Quote);
*/

let question1 = worksheet['C4'];

let option = worksheet['D6'];

let ans = worksheet['I4'];


console.log('Cell data->', question1.v, option.v, ans.v);
