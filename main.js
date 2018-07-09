const XLSX = require('xlsx')
const http = require('http');
const fs = require('fs');
var workbook ;
var sheet_name_list ;
var sheet;
var url ='http://tanab.com/peerEvaluation/result.xlsx';


var download = function(url, dest, cb) {
  let file = fs.createWriteStream(dest);
  let request = http.get(url, function(response) {
    response.pipe(file);
    file.on('finish', function() {
      file.close(cb);
    });
  });
}

function init()
{	
download(url,'D:\\work\\projects\\excelread\\result.xlsx',function(){
 workbook = XLSX.readFile('result.xlsx');
 sheet_name_list = workbook.SheetNames;
 sheet = workbook.Sheets[sheet_name_list[0]];	 
});
}


function get_nth_row(rowNo)
{
    let nthRow = [];
    let range = XLSX.utils.decode_range(sheet['!ref']);
    let C, R = rowNo?rowNo:0;
    for(C = range.s.c; C <= range.e.c; ++C) {
        let cell = sheet[XLSX.utils.encode_cell({c:C, r:R})] 
        let hdr = undefined;
        if(cell && cell.t) hdr = XLSX.utils.format_cell(cell);

        nthRow.push(hdr);
    }
    return nthRow;
	
}

init();

 
