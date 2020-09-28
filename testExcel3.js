var express = require('express');
var oracledb = require('oracledb');
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Worksheet Name');

var config = {
    user: process.env.USERNAME,
    password: process.env.PASSWORD,
    connectString: process.env.DATABASE
}

var app = express();
app.use(express.json());

app.post('/excel', function(req,res){
  //console.log('get in');
  res.send("hello");
   oracledb.getConnection(config, (err, conn) =>{
       todoWork(err, conn, req);
   });
});

app.listen(3000);
console.log('Server running at http://localhost:3000');

function todoWork(err, connection, req) {
    if (err) {
        console.error(err.message);
        return;
    }

    console.log(req.body.name);

    connection.execute(req.body.name, [], function (err, result) {
        if (err) {
            console.error(err.message);
            doRelease(connection);
            return;
        }
        //console.log(result.metaData);  //테이블 스키마
        //console.log(result.rows);  //데이터

        //엑셀다운로드
        const headingColumnNames = [
            "header1",
        ]

        //Write Column Title in Excel file
        let headingColumnIndex = 1;
        headingColumnNames.forEach(heading => {
            ws.cell(1, headingColumnIndex++)
                .string(heading)
        });

        //Write Data in Excel file
        let rowIndex = 2;
        result.rows.forEach( record => {
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
        wb.write('TeacherData.xlsx');
        //끝

        doRelease(connection);
    });
}

function doRelease(connection) {
    connection.release(function (err) {
        if (err) {
            console.error(err.message);
        }
    });
}
