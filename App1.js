const express = require('express');
const sql = require('mssql');
const ExcelJS = require('exceljs');
const fs = require('fs');
const moment = require('moment');

const app = express();
const cors = require('cors');

app.use(cors());
app.use(express.json());

// Database connection configuration
const config = {
  server: "WIN-0KU3OO5N3LT", // Replace with your server name
  port: 1433,
  user: 'sa', // Replace with your SQL Server username
  password: 'arka', // Replace with your SQL Server password
  database: 'opr4',
  options: {
    encrypt: true, // Enable encryption
    trustServerCertificate: true, // Accept self-signed certificates
    cryptoCredentialsDetails: {
      minVersion: 'TLSv1.2' // Set minimum TLS version for security
    }
  }
};

// Endpoint to import Excel data to MSSQL table
app.post('/import-excel', async (req, res) => {
  try {
    // Specify the path to the Excel file
    const filePath = 'C:/Users/ADMIN/Desktop/test.xlsx';

    // Check if the file exists
    if (!fs.existsSync(filePath)) {
      return res.status(400).json({ error: 'File not found' });
    }

    // Read the Excel file
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);

    // Connect to the database
    await sql.connect(config);

    // Iterate over rows in the Excel file
    for (let i = 2; i <= worksheet.rowCount; i++) { // Start from row 2 to skip header row
      const row = worksheet.getRow(i);
      const partno = row.getCell(1).text;
      const startdate = moment(row.getCell(2).text, "M/D/YYYY h:mm A").toISOString();
      const enddate = moment(row.getCell(3).text, "M/D/YYYY h:mm A").toISOString();
      const resource = row.getCell(4).text;
      const operationname = row.getCell(5).text; // New column: Operation Name

      // Check if the record already exists in the database
      const checkQuery = `
        SELECT COUNT(*) AS count
        FROM users
        WHERE partno = @partno
        AND startdate = @startdate
        AND enddate = @enddate
        AND resource = @resource
        AND operationname = @operationname
      `;

      const checkRequest = new sql.Request();
      checkRequest.input('partno', sql.NVarChar, partno);
      checkRequest.input('startdate', sql.DateTimeOffset, startdate);
      checkRequest.input('enddate', sql.DateTimeOffset, enddate);
      checkRequest.input('resource', sql.VarChar, resource);
      checkRequest.input('operationname', sql.VarChar, operationname);

      const { recordset } = await checkRequest.query(checkQuery);
      const existingCount = recordset[0].count;

      if (existingCount === 0) {
        // Define SQL query to insert data into the table
        const insertQuery = `
          INSERT INTO users (partno, startdate, enddate, resource, operationname)
          VALUES (@partno, @startdate, @enddate, @resource, @operationname)
        `;

        // Prepare SQL statement
        const insertRequest = new sql.Request();
        insertRequest.input('partno', sql.NVarChar, partno);
        insertRequest.input('startdate', sql.DateTimeOffset, startdate);
        insertRequest.input('enddate', sql.DateTimeOffset, enddate);
        insertRequest.input('resource', sql.VarChar, resource);
        insertRequest.input('operationname', sql.VarChar, operationname);

        // Execute SQL query to insert data
        await insertRequest.query(insertQuery);
      }
    }

    // Close SQL connection
    await sql.close();

    res.status(200).json({ message: 'New data imported successfully' });
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Start the server
const PORT = process.env.PORT || 3100;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});