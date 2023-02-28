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
        
        wb.xlsx.readFile(`./input/${file}`).then(() => {
            // const ws = wb.getWorksheet('Sheet0')
            const ws = wb.worksheets[0]; //the first one;

            let rowCount = ws.actualRowCount
            ws.getRow(rowCount + 2).getCell('A').value = 'Violation Report Parameters'
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

                }
                if(i === rowCount + 6){
                    ws.getRow(i).getCell('A').value = 'Driving Hours per Week'
                    ws.getRow(i).getCell('B').value = '56 (hours)'

                }
            }
            wb.clearThemes();
            wb.xlsx.writeFile(`./output/${file}`)
            console.log("Saving to output folder: " + file)
        })
        .catch(console.log)
        
        
    })
})

