const fs = require('fs')
const ExcelJS = require('exceljs');

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
        const wb = new ExcelJS.Workbook();
        let fontSize = 32
        let rowHeight = 100
        
        wb.xlsx.readFile(`./input/${file}`).then(() => {
            // const ws = wb.getWorksheet('Sheet0')
            const ws = wb.worksheets[0]; //the first one;
            ws.views = [{ zoomScale: 35, zoomScaleNormal: 35 }]
            ws.pageSetup.paperSize = 9
            ws.pageSetup.fitToPage = true
            ws.pageSetup.margins = {
                left: 0.7, right: 0.7,
                top: 0.75, bottom: 0.75,
                header: 0.3, footer: 0.3
            };

            for(let i = 1; i < 23; i++){
                ws.getColumn(i).width = 38;
            }
            let rowCount = ws.actualRowCount
            for(let i = 1; i <= rowCount; i++){
                ws.getRow(i).font = { name: 'Arial Narrow', size: fontSize }
                ws.getRow(i).height = rowHeight
                let row = ws.getRow(i)
                row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true}
                    cell.border = {
                        top: {style:'thin'},
                        left: {style:'thin'},
                        bottom: {style:'thin'},
                        right: {style:'thin'}
                    }
                });
            }
            ws.getRow(1).height = rowHeight + 20

            ws.getRow(rowCount + 1).height = rowHeight + 20
            ws.getRow(rowCount + 2).getCell('A').value = 'Violation Report Parameters'
            ws.getRow(rowCount + 2).getCell('A').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true}
            ws.getRow(rowCount + 2).font = { name: 'Arial Narrow', size: fontSize}
            for(let i = rowCount + 4; i <= rowCount + 6; i++){
                if(i === rowCount + 4){
                    ws.getRow(i).getCell('A').value = 'Idle Duration'
                    ws.getRow(i).getCell('B').value = '10 (min)'
                    ws.getRow(i).getCell('D').value = 'Stop Duration'
                    ws.getRow(i).getCell('E').value = '10 (min)'
                    ws.getRow(i).getCell('G').value = 'Over Speed on Highway'
                    ws.getRow(i).getCell('H').value = '60 (km/h)'
                    ws.getRow(i).getCell('J').value = 'Over Speed Duration on Highway'
                    ws.getRow(i).getCell('K').value = '10 (sec)'
                    ws.getRow(i).font = { name: 'Arial Narrow', size: fontSize }
                    ws.getRow(i).height = rowHeight + 20
                    let row = ws.getRow(i)
                    row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true}
                        cell.border = {
                            top: {style:'thin'},
                            left: {style:'thin'},
                            bottom: {style:'thin'},
                            right: {style:'thin'}
                        }
                    });
                }
                if(i === rowCount + 5){
                    ws.getRow(i).getCell('A').value = 'Over Speed Duration on Non Highway'
                    ws.getRow(i).getCell('B').value = '10 (sec)'
                    ws.getRow(i).getCell('D').value = 'Continuous Driving Hours'
                    ws.getRow(i).getCell('E').value = '4.5 (hours)'
                    ws.getRow(i).getCell('G').value = 'Working Hours'
                    ws.getRow(i).getCell('H').value = '12 (hours)'
                    ws.getRow(i).getCell('J').value = 'Rest Hours'
                    ws.getRow(i).getCell('K').value = '11 (hours)'
                    ws.getRow(i).font = { name: 'Arial Narrow', size: fontSize }
                    ws.getRow(i).height = rowHeight + 20
                    let row = ws.getRow(i)
                    row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true}
                        cell.border = {
                            top: {style:'thin'},
                            left: {style:'thin'},
                            bottom: {style:'thin'},
                            right: {style:'thin'}
                        }
                    });

                }
                if(i === rowCount + 6){
                    ws.getRow(i).getCell('A').value = 'Driving Hours per Week'
                    ws.getRow(i).getCell('B').value = '56 (hours)'
                    ws.getRow(i).font = { name: 'Arial Narrow', size: fontSize }
                    ws.getRow(i).height = rowHeight + 20
                    let row = ws.getRow(i)
                    row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true}
                        cell.border = {
                            top: {style:'thin'},
                            left: {style:'thin'},
                            bottom: {style:'thin'},
                            right: {style:'thin'}
                        }
                    });

                }
            }
            wb.clearThemes();
            wb.xlsx.writeFile(`./output/${file}`)
            console.log("Saving to output folder: " + file)
        })
        .catch(console.log)
        
        
    })
})

