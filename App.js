const sql = require('msnodesqlv8');
const express = require('express');
const app = express();
const cors = require('cors')
app.use(cors())
app.use(express.json())
const ExcelJS = require('exceljs');


// Database connection configuration
const connectString = "server=WIN-0KU3OO5N3LT\\SQLEXPRESS;Database=opr4;Trusted_Connection=yes;Driver={Sql Server Native Client 11.0}";

// Endpoint to execute SQL query
app.post('/hello', (req, res) => {
    const query = `USE opr4;
    WITH MaxRQ AS (
            SELECT MAX(RQ) AS MaxRQ FROM YourTableName12
        ),
        OrderedData AS (
            SELECT 
                RQ,
                VALUE,
                ROW_NUMBER() OVER (PARTITION BY RQ ORDER BY RQ) AS RowNum
            FROM [opr4].[RndSuite].[RndtRqIiSpreadSheetCell]
            JOIN MaxRQ ON RQ > MaxRQ.MaxRQ
        ),
        PivotedData AS (
            SELECT 
                RQ,
                MAX(CASE WHEN RowNum % 4 = 1 THEN VALUE END) AS OperationName,
                MAX(CASE WHEN RowNum % 4 = 2 THEN VALUE END) AS OperationNo,
                MAX(CASE WHEN RowNum % 4 = 3 THEN VALUE END) AS CycleTime,
                MAX(CASE WHEN RowNum % 4 = 0 THEN VALUE END) AS EquipmentType
            FROM OrderedData
            GROUP BY RQ, (RowNum - 1) / 4
        )
        INSERT INTO YourTableName12 (RQ, OperationName, OperationNo, CycleTime, EquipmentType)
        SELECT RQ, OperationName, OperationNo, CycleTime, EquipmentType
        FROM PivotedData
        WHERE OperationName <> 'OperationName'
            AND OperationNo <> 'OperationNo'
            AND CycleTime <> 'CycleTime'
            AND EquipmentType <> 'EquipmentType';
    
        USE opr4;
        INSERT INTO YourTableName123 (Ranking, RQ, SC, OperationName, OperationNo, CycleTime, EquipmentType)
        SELECT
            ROW_NUMBER() OVER (ORDER BY sc.SC) AS Ranking,
            tbl.RQ,
            sc.SC,
            tbl.OperationName,
            tbl.OperationNo,
            tbl.CycleTime,
            tbl.EquipmentType
        FROM
            [opr4].[RndSuite].[RndtRqSc] sc
        INNER JOIN
            [opr4].[dbo].[YourTableName12] tbl ON sc.RQ = tbl.RQ
        WHERE
            tbl.RQ > (SELECT MAX(RQ) FROM YourTableName123);
    
        USE opr4;
        INSERT INTO YourTablecheck (RQ, SC_VALUE,  OperationName, OperationNo, CycleTime, EquipmentType,PRIORITY)
        SELECT
            tbl.RQ,
            sc.SC_VALUE,
            tbl.OperationName,
            tbl.OperationNo,
            tbl.CycleTime,
            tbl.EquipmentType,
            sc.PRIORITY
        FROM
            [opr4].[RndSuite].[RndtSc] sc
        INNER JOIN
            [opr4].[dbo].[YourTableName123] tbl ON sc.SC = tbl.SC
        WHERE
            tbl.RQ > (SELECT MAX(RQ) FROM YourTablecheck)
            UPDATE [opr4].[dbo].[YourTablecheck]
        SET Batch = rq_ii.IIVALUE
        FROM [opr4].[dbo].[YourTablecheck] yt
        LEFT JOIN [opr4].[RndSuite].[RndtRqIi] rq_ii ON yt.RQ = rq_ii.RQ
        WHERE yt.Batch IS NULL
        AND rq_ii.IIVALUE IS NOT NULL;
;`;

    let responseSent = false; // Flag to track response

    sql.query(connectString, query, (err, rows) => {
        if (!responseSent) { // Check if response has been sent
            responseSent = true; // Set flag to true to indicate response has been sent
            if (err) {
                console.error('Error executing query:', err);
                return res.status(500).json({ error: 'Error executing query' });
            } else {
                console.log('Query executed successfully');
                return res.status(200).json({ message: 'Query executed successfully', data: rows });
            }
        }
    });
});



app.post('/hello1', (req, res) => {
    const query = `SELECT * FROM YourTablecheck`;

    sql.query(connectString, query, (err, rows) => {
        if (err) {
            console.error('Error executing query:', err);
            res.status(500).json({ error: 'Error executing query' });
        } else {
            console.log('Query executed successfully');

            // Export data to Excel
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Data');

            // Add headers
            const headers = Object.keys(rows[0]);
            worksheet.addRow(headers);

            // Add rows
            rows.forEach(row => {
                const rowData = Object.values(row);
                worksheet.addRow(rowData);
            });

            // Generate Excel file
            workbook.xlsx.writeBuffer()
                .then(buffer => {
                    // Set response headers for Excel download
                    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                    res.setHeader('Content-Disposition', 'attachment; filename=exported_data.xlsx');
                    res.setHeader('Content-Length', buffer.length);
                    res.send(buffer);
                })
                .catch(err => {
                    console.error('Error generating Excel:', err);
                    res.status(500).json({ error: 'Error generating Excel' });
                });
        }
    });
});

const PORT = process.env.PORT || 3100;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});