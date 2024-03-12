const express = require('express');
const sql = require('mssql');

// Create Express app
const app = express();
const PORT = process.env.PORT || 3200;
const cors = require('cors');

app.use(cors());
// MSSQL database configuration

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
  }

// Endpoint to fetch data from MSSQL table
// Backend code
// Backend code
app.get('/api/data/:sampleId', async (req, res) => {
  try {
      const sampleId = req.params.sampleId;

      // Connect to the database
      await sql.connect(config);

      // Query to fetch all details for the specified sample from the table
      const result = await sql.query(`SELECT * FROM [opr4].[dbo].[YourTable2] WHERE RQ_VALUE= '${sampleId}'`);

      // Close the connection
      await sql.close();

      // Send the data as JSON response
      res.json(result.recordset);
  } catch (error) {
      console.error('Error fetching data:', error.message);
      res.status(500).json({ error: 'Internal Server Error' });
  }
});



// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});