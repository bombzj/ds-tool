const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const BWIP = require('bwip-js');
const sharp = require('sharp'); // For image manipulation
const ProgressBar = require('progress');
const Jimp = require('jimp');
const { PDFDocument } = require('pdf-lib');


async function createBarcodePDF(barcodeImagePath, outputFilePath) {
    const image = await Jimp.read(barcodeImagePath);
    const imageBuffer = await image.getBufferAsync(Jimp.MIME_PNG);

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage();
    const imageEmbed = await pdfDoc.embedPng(imageBuffer);
    const { width, height } = image.bitmap;
    page.setSize(width, height)

    page.drawImage(imageEmbed, {
        x: 0,
        y: 0,
        width,
        height,
    });

    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputFilePath, pdfBytes);
    // remove file at barcodeImagePath
    fs.rmSync(barcodeImagePath)
}

const inputFolderPath = './';

const files = fs.readdirSync(inputFolderPath);

setTimeout(() => {
    console.log(code)
}, 999999999);

const currentDate = new Date();
const year = currentDate.getFullYear();
const month = String(currentDate.getMonth() + 1).padStart(2, '0');
const day = String(currentDate.getDate()).padStart(2, '0');
let font, font2

async function main() {
    font = await Jimp.loadFont(Jimp.FONT_SANS_64_BLACK)
    font2 = await Jimp.loadFont(Jimp.FONT_SANS_32_BLACK)
    const orderList = []
    let fileId = 1001;
    for (let file of files) {
        if (path.extname(file) === '.xlsx' && !path.basename(file).startsWith("~")) {
            const workbook = xlsx.readFile(path.join(inputFolderPath, file));
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
            let filenameFull = path.basename(file)
            const filenameparts = filenameFull.split('-')
            if(filenameparts.length == 2) {
                fileId = parseInt(filenameparts[1])
                filenameFull = filenameparts[0]
            }
            const basename = path.parse(filenameFull).name
            console.log("开始处理：" + file)

            const outputFolderPath = '.';
            // const outputFolderPath = './' + basename;
            const outputFolderOriginal = path.join(outputFolderPath, "原图")
            const outputPdfPath = path.join(outputFolderPath, '发货面单');
            // Create the output folder if it doesn't exist
            if (!fs.existsSync(outputFolderPath)) {
                fs.mkdirSync(outputFolderPath);
            }
            if (!fs.existsSync(outputFolderOriginal)) {
                fs.mkdirSync(outputFolderOriginal);
            }
            if (!fs.existsSync(outputPdfPath)) {
                fs.mkdirSync(outputPdfPath);
            }
            // check all rows, if pcs > 1, split this row into pcs rows, new rows added after original row
            for (let i = 0; i < rows.length; i++) {
                const [orderNumber, sku, code, imageUrl, productName, pcs, convertExt] = rows[i];
                if (parseInt(pcs) > 1) {
                    for (let j = 1; j < parseInt(pcs); j++) {
                        rows.splice(i + j, 0, [orderNumber, sku, code, imageUrl, productName, 1, convertExt])
                    }
                    rows[i][5] = 1
                }
            }

            const orderMap = new Map()  // set total pcs of each orderNumber in rows
            const codeMap = new Map() 
            for (let i = 0; i < rows.length; i++) {
                const [orderNumber, sku, code, imageUrl, productName, pcs, convertExt] = rows[i];
                if (orderMap.has(orderNumber)) {
                    const n = orderMap.get(orderNumber) + 1
                    orderMap.set(orderNumber, n)
                    // rows[i][0] += `-${n}`
                } else {
                    orderMap.set(orderNumber, 1)
                }
            }
            for (let i = 0; i < rows.length; i++) {
                rows[i][7] = orderMap.get(rows[i][0])
            }

            for (let i = 0; i < rows.length; i++) {
                const [orderNumber, sku, code, imageUrl, productName, pcs, convertExt, total] = rows[i];
                if (imageUrl === undefined || !imageUrl.startsWith("http")) continue
                
                const ext = imageUrl.split('.').pop().split('?')[0];
                if(convertExt == undefined || convertExt == "") {
                    convertExt = ext
                }
                const goodsSerial = `${basename}-${fileId}`
                const productPath = path.join(outputFolderOriginal, productName)
                if (!fs.existsSync(productPath)) {
                    fs.mkdirSync(productPath);
                }

                const downloadImage = async (url, retry = 5) => {
                    for (let i = 0; i < retry; i++) {
                        var bar = new ProgressBar(`${i == 0 ? "正在下载" : "重试下载"}：${orderNumber} [:bar] :result`, { total: 30 });
                        bar.update(0, { result: "" })
                        const response = await axios.get(url, { responseType: 'stream' }).catch(e => { })
                        if (!response) {
                            bar.render({ result: "失败" })
                            bar.terminate()
                            continue
                        }
                        const totalSize = parseInt(response.headers['content-length'], 10);
                        let downloadedSize = 0;
                        const dataBuffer = [];
                        let progress = 0
                        response.data.on('data', chunk => {
                            downloadedSize += chunk.length;
                            dataBuffer.push(chunk);
                            const updateProgress = Math.floor(downloadedSize * 30 / totalSize)
                            if (updateProgress > progress) {
                                if (updateProgress < 30)
                                    bar.tick(updateProgress - progress)
                                progress = updateProgress
                            }
                        });
                        await new Promise((resolve) => {
                            response.data.on('end', resolve);
                            response.data.on('error', resolve);
                        })
                        if (dataBuffer.length != 0) {
                            bar.update(1, { result: "成功" })
                            return Buffer.concat(dataBuffer, downloadedSize);
                        } else {
                            bar.render({ result: "失败" })
                            bar.terminate()
                        }
                    }
                }

                let orderData = codeMap.get(code)
                if (orderData) {
                    orderData.number++
                } else {
                    orderData = {
                        number: 1,
                        piece: 0,
                        orderNumber, sku, code
                    }
                    codeMap.set(code, orderData)
                }

                orderData.piece++

                const splitCode = total == 1 ? code : `${code}-${orderData.piece}`
                // const oriFile = path.join(outputFolderOriginal, productName, `${goodsSerial}.${convertExt}`)
                const oriFile = path.join(outputFolderOriginal, productName, `${splitCode}.${convertExt}`)     // 原始图保存文件

                const saveImageBarcode = async (data, cached = false) => {
                    const imageBuffer = data;
                    if (!cached)
                        fs.writeFileSync(oriFile, imageBuffer);
                    return await generateBarcode(code, [orderNumber, `Total: ${total} pcs`, currentDate.toLocaleString(), code], path.join(outputPdfPath, code))
                }

                let result = 0
                if (fs.existsSync(oriFile)) {
                    console.log("从缓存读取：" + orderNumber)
                    console.log(`开始生成pdf ${orderNumber}`)
                    result = await saveImageBarcode(fs.readFileSync(oriFile), true)
                    if(result == -2) {
                        console.log("图片已损坏，重新开始下载")
                    }
                }
                if(result != 1) {
                    const data = await downloadImage(imageUrl)
                    if (data) {
                        console.log(`开始生成pdf ${orderNumber}`)
                        if(ext != convertExt) {
                            const image = await sharp(data).toFormat(convertExt).toBuffer()
                            result = await saveImageBarcode(image)
                        }
                        else
                            result = await saveImageBarcode(data)
                        if(result == -2) {
                            console.log("下载的图片已损坏")
                        }
                    }
                }

                const orderDetail = [
                    code, orderNumber, goodsSerial, 10001
                ]
                orderDetail[7] = "条"
                orderDetail[15] = productName
                orderDetail[16] = splitCode
                orderList.push(orderDetail)
                fileId++;
            }
            console.log("文件处理完成：" + file)
        }
    }
    if(orderList.length == 0) {
        console.log("没有找到xlsx文件")
        return
    }

    const orderListFolder = "订单列表"
    if (!fs.existsSync(orderListFolder)) {
        fs.mkdirSync(orderListFolder);
    }

    {
        console.log("开始生成货品档案")
        const rowData = [
            "货品名称", "货品英文名称", "货品编号", "分类编号", "分类名称", "别名", "品牌", "单位", "辅助单位1", "转换率1", "辅助单位2", "转换率2", "辅助单位3", "转换率3", "常用单位", "规格", "条码"
        ];
    
    
        for (let i = 0; i < orderList.length; i++) {
            const row = orderList[i]
            // row[2] = `QHSTJ${year.toString().substring(2)}${month}${day}-${i + startTaskNumber}`
        }
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.aoa_to_sheet([rowData, ...orderList]);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        const defaultCol = { wch: 18 }
        const colConfig = worksheet['!cols'] = [
            defaultCol,
            defaultCol,
            defaultCol,
        ];
        colConfig[16] = defaultCol

        const xlsFile = `货品档案-${year}-${month}-${day}.xlsx`
        try {
            xlsx.writeFile(workbook, path.join(orderListFolder, xlsFile));
            console.log(`列表已生成 ${xlsFile}`)
        } catch (e) {
            console.log(`列表${xlsFile}无法覆盖，文件正在使用中`)
        }
    }

    {
        console.log("开始生成发货表")
        const rowData = [
            "货品编号", "名称", "订单", "物流单号", "图片", "产品数量"
        ];
    
        let orderList2 = []
        for (let i = 0; i < orderList.length; i++) {
            const row = orderList[i]
            orderList2.push([
                row[2], row[15], row[1], row[16], '', 1
            ])
        }
    
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.aoa_to_sheet([rowData, ...orderList2]);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        const defaultCol = { wch: 18 }
        const colConfig = worksheet['!cols'] = [
            defaultCol,
            defaultCol,
            defaultCol,
        ];

        const xlsFile = `发货表-${year}-${month}-${day}.xlsx`
        try {
            xlsx.writeFile(workbook, path.join(orderListFolder, xlsFile));
            console.log(`列表已生成 ${xlsFile}`)
        } catch (e) {
            console.log(`列表${xlsFile}无法覆盖，文件正在使用中`)
        }
    }
    
}



function generateBarcode(code, messages, filename) {
    return new Promise((resolve, reject) => {
        // if image already exists, then return
        if (fs.existsSync(filename + ".png")) {
            resolve(1)
            return
        }


        const barcodeOptions = {
            bcid: 'code128', // Barcode type
            text: code, // Text to encode
            scale: 4, // Barcode scaling factor
            height: 10, // Barcode height, in pixels
            includetext: true, // Show human-readable text below the barcode
        };

        // Create the barcode image
        BWIP.toBuffer(barcodeOptions, (err, png) => {
            if (err) throw err;

            // Load the barcode image
            Jimp.read(png)
                .then(barcodeImage => {
                    // Resize the barcode image to fit within the canvas
                    // barcodeImage.resize(700, 150);

                    // Create a blank canvas with size 800x800
                    new Jimp(800, 600, "#FFFFFF", (backgroundColorErr, canvas) => {
                        if (backgroundColorErr) throw backgroundColorErr;

                        // Merge the barcode image with the canvas by placing it in the center
                        const x = (canvas.bitmap.width - barcodeImage.bitmap.width) / 2;
                        const y = (canvas.bitmap.height - barcodeImage.bitmap.height) / 3;

                        canvas.composite(barcodeImage, x, y + 20);

                        // Add the number "15" at the bottom of the canvas
                        let nex = 50;
                        canvas.print(font, 50, 300 + nex, messages[0]);
                        canvas.print(font, 50, 380 + nex, messages[1]);
                        canvas.print(font2, 50, 460 + nex, messages[2]);
                        

                        canvas.print(font, 100, 60, messages[3]);

                        // Save the final image as a PNG file
                        canvas.write(filename + ".png", async (saveErr) => {
                            if (saveErr) console.log("条码生成出错: " + code);
                            await createBarcodePDF(filename + ".png", filename.replace('条码', '') + '.pdf')
                            resolve(1)
                        });
                    });
                })
                .catch(barcodeImageErr => {
                    console.error(barcodeImageErr);
                    resolve(-2)
                });
        });

    })
}

process.on('uncaughtException', UncaughtExceptionHandler);

function UncaughtExceptionHandler(err) {
    console.log("err: ", err);
    console.log("Stack trace: ", err.stack);
}

main()