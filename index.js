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
            const ws = wb.getWorksheet('Sheet0')
            let rowCount = ws.actualRowCount
            ws.getRow(rowCount + 2).getCell('A').value = 'Daily Violation Report Parameters'
            for(let i = rowCount + 4; i <= rowCount + 6; i++){
                if(i === rowCount + 4){
                    ws.getRow(i).getCell('A').value = 'Idle Duration'
                    ws.getRow(i).getCell('B').value = '10 (min)'
                    ws.getRow(i).getCell('D').value = 'Stop Duration'
                    ws.getRow(i).getCell('E').value = '10 (min)'
                    ws.getRow(i).getCell('G').value = 'Over Speed on Highway'
                    ws.getRow(i).getCell('H').value = '65 (km)'
                    ws.getRow(i).getCell('J').value = 'Over Speed on Non Highway'
                    ws.getRow(i).getCell('K').value = '65 (km)'

                }
                if(i === rowCount + 5){
                    ws.getRow(i).getCell('A').value = 'Over Speed Duration on Highway'
                    ws.getRow(i).getCell('B').value = '10 (sec)'
                    ws.getRow(i).getCell('D').value = 'Over Speed Duration on Non Highway'
                    ws.getRow(i).getCell('E').value = '10 (sec)'
                    ws.getRow(i).getCell('G').value = 'Continuous Driving Hours'
                    ws.getRow(i).getCell('H').value = '4 (hours)'
                    ws.getRow(i).getCell('J').value = 'Total Driving Hours'
                    ws.getRow(i).getCell('K').value = '9 (hours)'

                }
                if(i === rowCount + 6){
                    ws.getRow(i).getCell('A').value = 'Working Hours'
                    ws.getRow(i).getCell('B').value = '12 (hours)'
                    ws.getRow(i).getCell('D').value = 'Rest Hours'
                    ws.getRow(i).getCell('E').value = '10 (hours)'
                    ws.getRow(i).getCell('G').value = 'Driving Hours per Week'
                    ws.getRow(i).getCell('H').value = '54 (hours)'
                    ws.getRow(i).getCell('J').value = 'Work Hours per Week'
                    ws.getRow(i).getCell('K').value = '72 (hours)'

                }
            }
            wb.clearThemes();
            wb.xlsx.writeFile(`./output/${file}`)
            console.log("Saving to output folder: " + file)
        })
        .catch(console.log)
        
        
    })
})