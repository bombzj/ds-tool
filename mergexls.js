const { raw } = require('bwip-js');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// 读取配置表，根据配置表，读取店小秘表格，以及各个客户的表格，进行映射和合并，然后输出到一个新的表格中

async function doit() {
    // read from config.xlsx, header: folderName, column1Pos, column2Pos, column3Pos...
    const workbook = xlsx.readFile('config.xlsx');
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const configList = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    const [outputPrefix, ...outputColumnName] = configList[0];
    const [mappingPrefix, ...mappingColumnName] = configList[1];
    const mappingMapArray = [];
    for(let i = 0;i < mappingColumnName.length;i++)
        mappingMapArray.push(new Map());
    if(configList.length < 3) {
        console.log('数据缺失');
        return;
    }
    const files = fs.readdirSync(".").filter(file => file.endsWith('.xlsx'));
    const mappingFiles = files.filter(file => file.startsWith(mappingPrefix + '_'));
    // 店小秘导出模板
    for(let i = 0; i < mappingFiles.length; i++) {
        const workbook = xlsx.readFile(mappingFiles[i]);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        const columnNameToPos = new Map();
        for(let j = 0; j < data[0].length; j++)
            columnNameToPos.set(data[0][j], j);
        for(let j = 1; j < data.length; j++) {
            const row = [];
            for(let k = 0; k < mappingColumnName.length; k++) {
                const pos = columnNameToPos.get(mappingColumnName[k]);
                if(pos >= 0) {
                    row[k] = data[j][pos];
                    if(!mappingMapArray[k].has(row[k])) {
                        mappingMapArray[k].set(row[k], row);
                    }
                }
            }
        }
    }

    for(let i = 2;i < configList.length; i++) {
        const [rawPrefix, ...rawColumnName] = configList[i];
        const rawFiles = files.filter(file => file.startsWith(rawPrefix + '_'));
        for(const fileName of rawFiles) {
            const workbook = xlsx.readFile(fileName);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
            const columnNameToPos = new Map();
            for(let j = 0; j < data[0].length; j++)
                columnNameToPos.set(data[0][j], j);

            const rows = [];
            for(let j = 1; j < data.length; j++) {
                const row = {};
                for(let k = 0; k < outputColumnName.length; k++) {
                    if(isNaN(rawColumnName[k])) {
                        const pos = columnNameToPos.get(rawColumnName[k]);
                        if(pos >= 0) {
                            row[outputColumnName[k]] = data[j][pos];
                        }
                    } else {
                        const mappingPos = parseInt(rawColumnName[k]) - 1;
                        const pos = columnNameToPos.get(rawColumnName[mappingPos]);
                        const mappedData = mappingMapArray[mappingPos].get(data[j][pos]);
                        if(mappedData) {
                            row[outputColumnName[k]] = mappedData[k];
                        }
                    }
                }
                rows.push(row);
            }

            const outputFile = outputPrefix + fileName;
            const newWorkbook = xlsx.utils.book_new();
            const newWorksheet = xlsx.utils.json_to_sheet(rows, {header: outputColumnName});
            xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'result');
            xlsx.writeFile(newWorkbook, outputFile);
        
            console.log('输出到 ' + outputFile);
        }



    }

    // let result = [];
    // for(let i = 1; i < configList.length; i++) {
    //     const [folderName, ...columnPos] = configList[i];
    //     // iterate through all files in the folder
    //     const files = fs.readdirSync(folderName);
    //     for(let j = 0; j < files.length; j++) {
    //         if(!files[j].endsWith('.xlsx')) {
    //             continue;
    //         }
    //         const workbook = xlsx.readFile(path.join(folderName, files[j]));
    //         for(let sheet = 0; sheet < workbook.SheetNames.length; sheet++) {
    //             const worksheet = workbook.Sheets[workbook.SheetNames[sheet]];
    //             const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    
    //             for(let k = 1; k < data.length; k++) {
    //                 if(data[k][1] === undefined || data[k][1] === '') {
    //                     continue;
    //                 }
    //                 let row = {};
    //                 for(let l = 0; l < columnPos.length; l++) {
    //                     row[columnName[l]] = data[k][columnPos[l] - 1];
    //                 }
    //                 result.push(row);
    //             }
    //         }
    //     }
    //     console.log(`folder ${folderName} done`);
    // }

}


doit();
// pause the console
process.stdin.resume();
process.stdin.on('data', function(data) {
    process.exit();
});

