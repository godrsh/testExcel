var express = require('express');
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Worksheet Name');

var app = express();
app.use(express.json());
//app.use(bodyParser().json());

console.log('router before');  //데이터
app.post('/excel', function(err,req,res){
   if(err){
       console.log("err");
   }

  console.log('write start');  //데이터

  const body = req.body;

  console.log(req.body);

  //엑셀다운로드
  const headingColumnNames = [
      "header_1"
  ]

  //Write Column Title in Excel file
  let headingColumnIndex = 1;
  headingColumnNames.forEach(heading => {
      ws.cell(1, headingColumnIndex++)
          .string(heading)
  });

  //Write Data in Excel file
  let rowIndex = 2;
  req.body.forEach( record => {
      let columnIndex = 1;

      Object.keys(record ).forEach(columnName =>{
        if(typeof(record [columnName]) == "string"){
          ws.cell(rowIndex,columnIndex++)
              .string(record [columnName] || '');
        }else if(typeof(record [columnName]) == "number"){
          ws.cell(rowIndex,columnIndex++)
              .number(record [columnName] || 0);
        }else{
          ws.cell(rowIndex,columnIndex++)
              .string('');
        }
    });
      rowIndex++;
  });
  console.log('cell mapping finish');  //데이터

  wb.write('excelExport.xlsx');
  //끝
  console.log('write finish');  //데이터

  res.send("###################### excel make finish ####################");
  res.status(200).end();
});

app.listen(3000);
console.log('Server running at http://localhost:3000');
