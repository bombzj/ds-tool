const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');


async function doit() {
    // read from config.xlsx, header: folderName, column1Pos, column2Pos, column3Pos...
    const workbook = xlsx.readFile('config.xlsx');
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const configList = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    const [outputFile, ...columnName] = configList[0];

    let result = [];
    for(let i = 1; i < configList.length; i++) {
        const [folderName, ...columnPos] = configList[i];
        // iterate through all files in the folder
        const files = fs.readdirSync(folderName);
        for(let j = 0; j < files.length; j++) {
            if(!files[j].endsWith('.xlsx')) {
                continue;
            }
            const workbook = xlsx.readFile(path.join(folderName, files[j]));
            for(let sheet = 0; sheet < workbook.SheetNames.length; sheet++) {
                const worksheet = workbook.Sheets[workbook.SheetNames[sheet]];
                const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    
                for(let k = 1; k < data.length; k++) {
                    if(data[k][1] === undefined || data[k][1] === '') {
                        continue;
                    }
                    let row = {};
                    for(let l = 0; l < columnPos.length; l++) {
                        row[columnName[l]] = data[k][columnPos[l] - 1];
                    }
                    result.push(row);
                }
            }
        }
        console.log(`folder ${folderName} done`);
    }

    const newWorkbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.json_to_sheet(result, {header: columnName});
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'result');
    xlsx.writeFile(newWorkbook, outputFile);

    console.log('saved to ' + outputFile);
}


doit();
// pause the console
process.stdin.resume();
process.stdin.on('data', function(data) {
    process.exit();
});

