// Test on 127.0.0.1:10000

const express = require("express")
const app = express()

app.use(express.json())

app.post("/", (request, response) => {
    const Excel = require('exceljs');
    const wb = new Excel.Workbook();

    // File name of the Excel file
    const fileName = "scholarships.xlsx";
    // Keyword that is searched by user
    const { keyword } = request.body;

    // 2D array that contains each row where scholarship name matches keyword
    var scholarships = [];

    // Open the excel file
    wb.xlsx.readFile(fileName).then(() => {
        const ws = wb.getWorksheet('Sheet1'); // Read Sheet1
        const names = ws.getColumn(1);

        // For each cell in column 1...
        names.eachCell(function (cell, cellNumber) {
            // Add scholarship object to list if scholarship name matches with keyword
            if (cellNumber > 1 && cell.text.toLowerCase().indexOf(keyword.toLowerCase()) != -1) {
                let info = ws.getRow(cellNumber).values.slice(1);
                let scholarship_obj = {
                    name: info[0],
                    link: info[1],
                    categories: info[2],
                    amount: info[3],
                    deadline: info[4],
                    description: info[5]
                };
                scholarships.push(scholarship_obj);
            }
        });

        // Sends an array of scholarship objects as a response
        // Frontend uses this information to display the scholarships
        response.send({ scholarships });
    }).catch(err => {
        console.log(err.message);
    });
});

app.get("/", function (request, response) {
    response.sendFile(__dirname + "/index.html");
})

app.listen(10000, function () {
    console.log("Started application on port %d", 10000)
});
