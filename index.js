// //  The update table button, 
// after clicking it, stores all the MS SQL queries from the backend to the database table.

//****************************************************************************************/

self.addNewButton = function () {
    actionBar.createButtonInDynamicArea({
        id: 'NewButton',
        options: [{
            text: 'Update Table',
            iconClass: 'icon-cmdCheckmark24',
            action: function () {
                // Fetch data from the backend API
                fetch('http://localhost:3000/hello', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({}) // If you need to send any data in the request body
                })
                .then(response => response.json())
                .then(data => {
                    // Handle the response data
                    console.log(data); // For example, log the data to the console
                    // Add your custom logic here based on the response
                })
                .catch(error => {
                    // Handle any errors that occur during the fetch
                    console.error('Error fetching data:', error);
                });
            },
            visible: ko.computed(function () {
                // Define the visibility logic for the new button
                // You can customize this based on your requirements
                return true; // For example, always visible
            })
        }]
    });
};


// The update sample report button, after clicking it, stores all the MS SQL queries from the backend 
// to the database table.
        self.addNewButton5 = function () {
            actionBar.createButtonInDynamicArea({
                id: 'NewButton5',
                options: [{
                    text: 'Update Sample Report',
                    iconClass: 'icon-cmdAccept24',
                    action: function () {
                        // Fetch data from the backend API
                        fetch('http://localhost:3000/hello5', {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify({}) // If you need to send any data in the request body
                        })
                        .then(response => response.json())
                        .then(data => {
                            // Handle the response data
                            console.log(data); // For example, log the data to the console
                            // Add your custom logic here based on the response
                        })
                        .catch(error => {
                            // Handle any errors that occur during the fetch
                            console.error('Error fetching data:', error);
                        });
                    },
                    visible: ko.computed(function () {
                        // Define the visibility logic for the new button
                        // You can customize this based on your requirements
                        return true; // For example, always visible
                    })
                }]
            });
        };
//  *************************************************************************************
app.post('/hello5', (req, res) => {
    const query = `
        INSERT INTO YourTable1 (SC, SC_VALUE, PA_SHORT_DESC, PG, VALUE_F)
        SELECT
            [opr4].[RndSuite].[RndtSc].[SC],
            [opr4].[RndSuite].[RndtSc].[SC_VALUE],
            [opr4].[RndSuite].[RndtScPa].[PA_SHORT_DESC],
            [opr4].[RndSuite].[RndtScPa].[PG],
            [opr4].[RndSuite].[RndtScPa].[VALUE_F]
        FROM
            [opr4].[RndSuite].[RndtSc]
        INNER JOIN
            [opr4].[RndSuite].[RndtScPa] ON [opr4].[RndSuite].[RndtSc].[SC] = [opr4].[RndSuite].[RndtScPa].[SC] 
        WHERE
            [opr4].[RndSuite].[RndtSc].[SC] > (
                SELECT MAX(SC) FROM [opr4].[dbo].[YourTable1]
            );

        INSERT INTO YourTable2 (SC, SC_VALUE, PA_SHORT_DESC, LOW_SPEC, VALUE_F, HIGH_SPEC)
        SELECT
            t1.SC,
            t1.SC_VALUE,
            t1.PA_SHORT_DESC,
            s.LOW_SPEC,
            t1.VALUE_F,
            s.HIGH_SPEC
        FROM
            opr4.dbo.YourTable1 AS t1
        INNER JOIN
            opr4.RndSuite.RndtScPaSpa AS s ON t1.PG = s.PG 
                AND t1.SC = s.SC
        WHERE
            NOT EXISTS (
                SELECT 1
                FROM YourTable2 AS t2
                WHERE t1.SC_VALUE = t2.SC_VALUE
            );

        UPDATE YourTable2
        SET RESULT = 
            CASE 
                WHEN VALUE_F >= LOW_SPEC AND VALUE_F <= HIGH_SPEC THEN 'passed'
                ELSE 'failed'
            END;

        UPDATE [opr4].[dbo].[YourTable2]
        SET 
            RQ = (SELECT RQ FROM [opr4].[RndSuite].[RndtRqSc] WHERE [opr4].[RndSuite].[RndtRqSc].[SC] = [opr4].[dbo].[YourTable2].[SC]);

        UPDATE [opr4].[dbo].[YourTable2]
        SET 
            RQ_VALUE = (SELECT RQ_VALUE FROM [opr4].[RndSuite].[RndtRq] WHERE [opr4].[RndSuite].[RndtRq].[RQ] = [opr4].[dbo].[YourTable2].[RQ]);
    `;

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





// To import excel data to the ms sql table
 self.addNewButton3 = function () {
            actionBar.createButtonInDynamicArea({
                id: 'NewButton3',
                options: [{
                    text: 'Excel import',
                    iconClass: 'icon-cmdWorksheetAssign24',
                    action: function () {
                        // Make HTTP POST request to trigger the Excel import API
                        fetch('http://localhost:3100/import-excel', {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify({}) // Empty body as no data is required for the import
                        })
                        .then(response => {
                            if (response.ok) {
                                // Import successful
                                alert('Excel import successful');
                            } else {
                                // Import failed
                                alert('Excel import failed');
                            }
                        })
                        .catch(error => {
                            console.error('Error:', error);
                            alert('Internal server error');
                        });
                    },
                    visible: ko.computed(function () {
                        // Define the visibility logic for the new button
                        // You can customize this based on your requirements
                        return true; // For example, always visible
                    })
                }]
            });
        };



        // export database table data to the excel sheet
self.addNewButton1 = function () {
            actionBar.createButtonInDynamicArea({
                id: 'NewButton1',
                options: [{
                    text: 'Excel Export',
                    iconClass: 'icon-cmdWorksheetAssign24',
                    action: function () {
                        // Make a request to the backend endpoint
                        fetch('http://localhost:3000/hello1', { // Changed endpoint to match backend
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify({}) // Pass any necessary data here
                        })
                        .then(response => {
                            // Check if the response is successful
                            if (!response.ok) {
                                throw new Error('Network response was not ok');
                            }
                            // Convert the response to blob
                            return response.blob();
                        })
                        .then(blob => {
                            // Create a link element
                            const url = window.URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url;
                            a.download = 'exported_data.xlsx'; // Specify the filename
                            // Programmatically trigger a click event to start the download
                            document.body.appendChild(a);
                            a.click();
                            // Clean up
                            window.URL.revokeObjectURL(url);
                            document.body.removeChild(a);
                        })
                        .catch(error => {
                            console.error('Error:', error);
                            // Handle error
                        });
                    },
                    visible: ko.computed(function () {
                        // Define the visibility logic for the new button
                        // You can customize this based on your requirements
                        return true; // For example, always visible
                    })
                }]
            });
        };


        // generate report based on request
self.addNewButton4 = function () {
            actionBar.createButtonInDynamicArea({
                id: 'NewButton4',
                options: [{
                    text: 'Sample Report',
                    iconClass: 'icon-cmdCheckmark24',
                    action: function () {
                        // Prompt the user to input the sample ID
                        var sampleId = prompt("Enter Request ID:");
        
                        // Make an HTTP request to fetch data for the selected sample from backend API
                        fetch(`http://localhost:3200/api/data/${sampleId}`)
                            .then(response => response.json())
                            .then(data => {
                                // Create table HTML dynamically using fetched data for the selected sample
                                var tableHtml = `
                                    <!DOCTYPE html>
                                    <html>
                                        <head>
                                            <title>Sample Report</title>
                                            <style>
                                                /* CSS styles */
                                                body {
                                                    display: flex;
                                                    justify-content: center;
                                                    align-items: center;
                                                    height: 100vh;
                                                }
                                                .center {
                                                    text-align: center;
                                                }
                                                .logo {
                                                    max-width: 80%;
                                                    height: auto;
                                                    margin-bottom: 20px;
                                                }
                                                table {
                                                    border-collapse: collapse;
                                                    width: 80%;
                                                    margin-bottom: 20px;
                                                }
                                                th, td {
                                                    border: 1px solid black;
                                                    padding: 8px;
                                                    text-align: left;
                                                }
                                                th {
                                                    background-color: #f2f2f2;
                                                }
                                                .passed {
                                                    color: green;
                                                }
                                                .failed {
                                                    color: red;
                                                }
                                            </style>
                                        </head>
                                        <body>
                                            <div class="center">
                                                <img class="logo" src="file:///C:/Users/ADMIN/Downloads/bavistech2_logo.jpg" alt="Logo">
                                                <table>
                                                    <thead>
                                                        <tr>
                                                            <th>RQ_VALUE</th>
                                                            <th>SC_VALUE</th>
                                                            <th>PA_SHORT_DESC</th>
                                                            <th>LOW_SPEC</th>
                                                            <th>VALUE_F</th>
                                                            <th>HIGH_SPEC</th>
                                                            <th>RESULT</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>`;
        
                                // Populate table rows with fetched data for the selected sample
                                data.forEach(row => {
                                    var resultClass = row.RESULT.toLowerCase() === 'passed' ? 'passed' : 'failed';
                                    tableHtml += `<tr>
                                        <td>${row.RQ_VALUE}</td>
                                        <td>${row.SC_VALUE}</td>
                                        <td>${row.PA_SHORT_DESC}</td>
                                        <td>${row.LOW_SPEC}</td>
                                        <td>${row.VALUE_F}</td>
                                        <td>${row.HIGH_SPEC}</td>
                                        <td class="${resultClass}">${row.RESULT}</td>
                                    </tr>`;
                                });
        
                                tableHtml += `
                                        </tbody>
                                    </table>
                                </div>
                            </body>
                        </html>`;
        
                        // Convert the HTML string to a Blob
                        var blob = new Blob([tableHtml], { type: 'text/html' });
        
                        // Create a download link
                        var link = document.createElement('a');
                        link.href = window.URL.createObjectURL(blob);
                        link.download = 'sample_report.html';
        
                        // Append the download link to the document body
                        document.body.appendChild(link);
        
                        // Trigger the download
                        link.click();
                    })
                    .catch(error => console.error('Error fetching data:', error));
            },
            visible: ko.computed(function () {
                return true; // Button always visible
            })
        }]
        });
        };

