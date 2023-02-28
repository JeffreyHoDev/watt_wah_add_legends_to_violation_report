const fs = require('fs')
const ExcelJS = require('exceljs');
const excelToJson = require('convert-excel-to-json');
var pdf = require('html-pdf');
var options = { format:"A4", orientation: "landscape", border: "1cm" };

if(!fs.existsSync('./input')){
    fs.mkdirSync('./input')
    console.log("Input folder created, you can now put excel files inside")
}

if(!fs.existsSync('./output')){
    fs.mkdirSync('./output')
    console.log("Output folder created, output result excel files will be gathered here")
}

fs.readdir('./input', (error, files) => {
    if (error) console.log(error)

    files.forEach(file => {
        let result = excelToJson({
            sourceFile: `./input/${file}`,
            header: {
                rows: 1
            }
        })

        let filenameIndex = file.indexOf('.')
        let filename = file.substr(0, filenameIndex)
        
        let html = `
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta http-equiv="X-UA-Compatible" content="IE=edge">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Document</title>
                <style>
                    #table {
                        font-family: Arial, Helvetica, sans-serif;
                        border-collapse: collapse;
                        width: 100%;
                        margin: .2rem;
                        font-size: 10px;
                    }

                    #table td, #table th {
                        border: 1px solid #ddd;
                        padding: 8px;
                    }

                    #table tr:nth-child(even){background-color: #f2f2f2;}

                    #table tr:hover {background-color: #ddd;}

                    #table th {
                        padding-top: 12px;
                        padding-bottom: 12px;
                        text-align: left;
                        background-color: #337DFF;
                        color: white;
                    }

                    #params-table {
                        font-family: Arial, Helvetica, sans-serif;
                        border-collapse: collapse;
                        width: 100%;
                        margin: .2rem;
                        font-size: 10px;
                    }

                    #params-table td {
                        border: 1px solid #ddd;
                        padding: 8px;
                    }
                </style>
            </head>
            <body>
                <h2>Violation Report Content</h2>
                <table id="table">
                    <tr>
                        <th>Plate No</th>
                        <th>Date</th>
                        <th>High Speed Alarm</th>
                        <th>Driver Fatigue</th>
                        <th>Phone Detection</th>
                        <th>Smoking Detection</th>
                        <th>Driver Distraction</th>
                        <th>Lane Departure</th>
                        <th>Forward Collision Warning</th>
                        <th>Following Distance Monitoring</th>
                        <th>Pedestrian Collision Warning</th>
                        <th>Yawning Detection</th>
                        <th>Idling alarm</th>
                        <th>Harsh Cornering</th>
                        <th>Harsh Acceleration</th>
                        <th>Harsh Braking</th>
                        <th>Unauthorized Parking</th>
                        <th>Seatbelt</th>
                        <th>Continuous Driving</th>
                        <th>Total Violation</th>
                        <th>Total Driving Hour</th>
                        <th>Total Working Hour</th>
                    </tr>
                    ${
                        result.Sheet0.map((item) => {
                            return `
                                <tr>
                                    <td>${item.A}</td>
                                    <td>${item.B}</td>
                                    <td>${item.C}</td>
                                    <td>${item.D}</td>
                                    <td>${item.E}</td>
                                    <td>${item.F}</td>
                                    <td>${item.G}</td>
                                    <td>${item.H}</td>
                                    <td>${item.I}</td>
                                    <td>${item.J}</td>
                                    <td>${item.K}</td>
                                    <td>${item.L}</td>
                                    <td>${item.M}</td>
                                    <td>${item.N}</td>
                                    <td>${item.O}</td>
                                    <td>${item.P}</td>
                                    <td>${item.Q}</td>
                                    <td>${item.R}</td>
                                    <td>${item.S}</td>
                                    <td>${item.T}</td>
                                    <td>${item.U}</td>
                                    <td>${item.V}</td>
                                </tr>
                            `
                        }).join('')
                    }
                </table>
                <h3>Violation Report Parameters</h3>
                <table id="params-table">
                    <tr>
                        <td>Idle Duration: 10(min)</td>
                        <td>Stop Duration: 10(min)</td>
                        <td>Over Speed on Highway: 60(km/h)</td>
                        <td>Over Speed Duration on Highway: 10(sec)</td>
                    </tr>
                    <tr>
                        <td>Over Speed Duration on non highway: 10(sec)</td>
                        <td>Continuous Driving Hours: 4.5(hours)</td>
                        <td>Working Hours: 12(hours)</td>
                        <td>Rest Hours: 11(hours)</td>
                    </tr>
                    <tr>
                        <td>Driving Hours per Week: 56(hours)</td>
                    </tr>
                </table>
            </body>
            </html>
        `



        pdf.create(html, options).toFile(`./output/${filename}.pdf`, function (err, result) {
            if (err) {
                return res.status(400).send({
                    message: errorHandler.getErrorMessage(err)
                });
            }
        })
        
    })
})
