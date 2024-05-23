const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const BWIP = require('bwip-js');
const sharp = require('sharp'); // For image manipulation
const ProgressBar = require('progress');
const Jimp = require('jimp');
const { PDFDocument } = require('pdf-lib');

let startTaskNumber = 8001
let currentTaskNumber = startTaskNumber


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
    let lastSplitTask
    for (let file of files) {
        if (path.extname(file) === '.xlsx' && !path.basename(file).startsWith("~")) {
            const workbook = xlsx.readFile(path.join(inputFolderPath, file));
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

            let basefilename = path.basename(file, '.xlsx')
            let splitfilename = basefilename.indexOf("--")
            if(splitfilename != -1)
                currentTaskNumber = startTaskNumber = parseInt(basefilename.substring(splitfilename + 2))

            const basename = path.parse(path.basename(file)).name
            console.log("开始处理：" + file)

            const outputFolderPath = './' + basename;
            const outputFolderOriginal = outputFolderPath + "原始图"
            // Create the output folder if it doesn't exist
            if (!fs.existsSync(outputFolderPath)) {
                fs.mkdirSync(outputFolderPath);
            }
            if (!fs.existsSync(outputFolderOriginal)) {
                fs.mkdirSync(outputFolderOriginal);
            }

            const orderMap = new Map()  // 原始图片

            for (let i = 0; i < rows.length; i++) {
                const [orderNumber, sku, color, pcs, code, imageUrl] = rows[i];
                if (imageUrl === undefined || !imageUrl.startsWith("http")) continue
                
                const ext = imageUrl.split('.').pop().split('?')[0];
                const oriFile = path.join(outputFolderOriginal, `${orderNumber}-${color}.${ext}`)  // 原始图保存文件

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


                const splitImage = async (data, cached = false) => {
                    const imageBuffer = data;
                    if (!cached)
                        fs.writeFileSync(oriFile, imageBuffer);
                    await generateBarcode(code, [orderNumber, `Total: ${pcs} pcs`, currentDate.toLocaleString()], path.join(outputFolderPath, orderNumber), oriFile)
                    return 1
                }

                let result = 0
                if (fs.existsSync(oriFile)) {
                    console.log("从缓存读取：" + orderNumber)
                    console.log(`开始生成pdf ${orderNumber}`)
                    result = await splitImage(fs.readFileSync(oriFile), true)
                    if(result == -2) {
                        console.log("图片已损坏，重新开始下载")
                    }
                }
                if(result != 1) {
                    const data = await downloadImage(imageUrl)
                    if (data) {
                        console.log(`开始生成pdf ${orderNumber}`)
                        result = await splitImage(data)
                        if(result == -2) {
                            console.log("下载的图片已损坏")
                        }
                    }
                }
            }
        }
    }
}



function generateBarcode(code, messages, filename, imageFile) {
    return new Promise((resolve, reject) => {

        const barcodeOptions = {
            bcid: 'code128', // Barcode type
            text: code, // Text to encode
            scale: 3, // Barcode scaling factor
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
                    // barcodeImage.resize(600, 120);

                    // Create a blank canvas with size 800x800
                    new Jimp(800, 800, "#FFFFFF", (backgroundColorErr, canvas) => {
                        if (backgroundColorErr) throw backgroundColorErr;

                        // Merge the barcode image with the canvas by placing it in the center
                        const x = (canvas.bitmap.width - barcodeImage.bitmap.width) / 2;
                        const y = (canvas.bitmap.height - barcodeImage.bitmap.height) / 3;

                        canvas.composite(barcodeImage, x, y + 200);

                        // Add the number "15" at the bottom of the canvas
                        let nex = 50;

                        canvas.print(font, 100, 500 + nex, messages[0]);
                        canvas.print(font, 100, 580 + nex, messages[1]);
                        canvas.print(font2, 100, 660 + nex, messages[2]);

                        // add imageFile at the top of the canvas
                        Jimp.read(imageFile).then(image => {
                            image.resize(700, 340)
                            canvas.composite(image, 50, 50)

                            // Save the final image as a PNG file
                            canvas.write(filename + ".png", (saveErr) => {
                                if (saveErr) console.log("条码生成出错: " + code);
                                createBarcodePDF(filename + ".png", filename.replace('条码', '') + '.pdf')
                                resolve(1)
                            });
                        }).catch(e => {
                            console.log("图片加载失败：" + imageFile)
                        })
                    });
                })
                .catch(barcodeImageErr => {
                    console.error(barcodeImageErr);
                });
        });

    })
s
}

process.on('uncaughtException', UncaughtExceptionHandler);

function UncaughtExceptionHandler(err) {
    console.log("err: ", err);
    console.log("Stack trace: ", err.stack);
}

main()