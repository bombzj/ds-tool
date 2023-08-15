const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const sharp = require('sharp'); // For image manipulation

const inputFolderPath = './';


const files = fs.readdirSync(inputFolderPath);

files.forEach(async file => {
  if (path.extname(file) === '.xlsx' && !path.basename(file).startsWith("~")) {
    const workbook = xlsx.readFile(path.join(inputFolderPath, file));
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    const basename = path.parse(path.basename(file)).name
    console.log("开始处理：" + basename)

    const outputFolderPath = './' + basename;
    const outputFolderOriginal = outputFolderPath + "原始图"
    // Create the output folder if it doesn't exist
    if (!fs.existsSync(outputFolderPath)) {
      fs.mkdirSync(outputFolderPath);
    }
    if (!fs.existsSync(outputFolderOriginal)) {
      fs.mkdirSync(outputFolderOriginal);
    }

    const codeMap = new Map()
    const codeMap2 = new Map()  // 原始图片

    for (let i = 1; i < rows.length; i++) {
      const [, sku ,code , imageUrl] = rows[i];
      const serialNumbers = [0, 1, 2, 3, 4];
      let n = 1;
      const regex = /-(\d+)P$/;
      const match = sku.match(regex);
      if(match){
        n = parseInt(match[1])
      }

      let serial2 = codeMap2.get(code)
      if(serial2){
        serial2++
      } else {
        serial2 = 1
      }
      codeMap2.set(code, serial2)
      const oriFile = path.join(outputFolderOriginal, `${code}-${serial2}.png`)  // 原始图保存文件

      const splitImage = (data) => {
        const imageBuffer = data;
        fs.writeFileSync(oriFile, imageBuffer);
        return Promise.all(serialNumbers.map(serialNumber => {
          return sharp(imageBuffer).metadata().then(metadata => {
              if(serialNumber >= n) return
              let serial = codeMap.get(code)
              if(serial){
                  serial++
              } else {
                  serial = 1
              }
              codeMap.set(code, serial)
              const width = Math.floor(metadata.width / 5)
          sharp(imageBuffer).extract({ left: width * serialNumber, top: 0, width, height: metadata.height })
            .trim()
            .toFile(path.join(outputFolderPath, `${code}-${serial}.png`));
          })
        }));
      }

      if(fs.existsSync(oriFile)) {
        console.log("从缓存读取：" + code)
        splitImage(fs.readFileSync(oriFile))
      } else {
        console.log("开始下载：" + code)
        await axios.get(imageUrl, { responseType: 'arraybuffer' })
          .then(response => {
              splitImage(response.data)
          })
          .then(() => {
            console.log(`Images for code ${code} downloaded and processed.`);
          })
          .catch(error => {
            console.error(`Error processing images for code ${code}:`, error);
          });
      }
    }
  }
});
