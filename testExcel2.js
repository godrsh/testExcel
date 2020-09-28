var oracledb = require('oracledb');
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Worksheet Name');

var config = {
    user: process.env.USERNAME,
    password: process.env.PASSWORD,
    connectString: process.env.DATABASE
}

oracledb.getConnection(config, (err, conn) =>{
    todoWork(err, conn);
});

function todoWork(err, connection) {
    if (err) {
        console.error(err.message);
        return;
    }
    connection.execute("sql입력자리", [], function (err, result) {
        if (err) {
            console.error(err.message);
            doRelease(connection);
            return;
        }

        //엑셀다운로드
        const headingColumnNames = [
            "SKU_CODE",
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
        wb.write('COSKD_BIZTP_TEST.xlsx');
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
