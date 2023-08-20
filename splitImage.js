const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const BWIP = require('bwip-js');
const sharp = require('sharp'); // For image manipulation
const ProgressBar = require('progress');
const Jimp = require('jimp');

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
                const [orderNumber, sku, code, imageUrl] = rows[i];
                if (imageUrl === undefined || !imageUrl.startsWith("http")) continue
                let n = 1;
                const regex = /-(\d+)P$/;
                const match = sku.match(regex);
                if (match) {
                    n = parseInt(match[1])
                }

                let orderData = orderMap.get(code)
                if (orderData) {
                    orderData.number++
                } else {
                    orderData = {
                        number: 1,
                        piece: 0,
                        orderNumber, sku, code
                    }
                    orderMap.set(code, orderData)
                }
                const oriFile = path.join(outputFolderOriginal, `${code}-${orderData.number}.png`)  // 原始图保存文件

                const downloadImage = async (url, retry = 5) => {
                    for (let i = 0; i < retry; i++) {
                        var bar = new ProgressBar(`${i == 0 ? "正在下载" : "重试下载"}：${code} [:bar] :result`, { total: 30 });
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
                    const imageSharp = sharp(imageBuffer)
                    let info
                    try {
                        info = await imageSharp
                        .raw()
                        .toBuffer({ resolveWithObject: true });
                    } catch(e) {
                        return -2
                    }
                    const metadata = info.info

                    const pixelArray = new Uint8ClampedArray(info.data.buffer);
                    const boundingList = []
                    let bounding = {
                        left: -1
                    }
                    const baseline = (metadata.width * Math.floor(metadata.height / 2)) << 2
                    for (let i = 0; i < metadata.width; i++) {
                        let transparent = pixelArray[baseline + i * 4 + 3] == 0
                        if (bounding.left == -1) {
                            if (!transparent) {
                                bounding.left = i
                            }
                        } else {
                            if (transparent) {
                                bounding.right = i
                                boundingList.push(bounding)
                                bounding = {
                                    left: -1
                                }
                            }
                        }
                    }
                    if (bounding.left != -1) {
                        bounding.right = metadata.width
                        boundingList.push(bounding)
                    }
                    if(boundingList.length != n){
                        console.log(`识别到${boundingList.length}个图片，与订单要求的${n}不一致`)
                    }

                    for (let i = 0; i < boundingList.length; i++) {
                        let bound = boundingList[i]
                        if (i >= n) break

                        orderData.piece++

                        const splitCode = `${code}-${orderData.piece}`
                        const outputFile = path.join(outputFolderPath, `${splitCode}.png`)
                        if (!fs.existsSync(outputFile)) {
                            const width = bound.right - bound.left// Math.floor(metadata.width / 5)
                            await sharp(imageBuffer).extract({ left: bound.left, top: 0, width, height: metadata.height })
                                .toFile(outputFile);
                        }
                        const orderDetail = [
                            code, orderNumber, , 5004
                        ]
                        orderDetail[7] = "条"
                        orderDetail[16] = splitCode
                        orderDetail[28] = "是"
                        orderList.push(orderDetail)
                    }
                    return 1
                }

                let result = 0
                if (fs.existsSync(oriFile)) {
                    console.log("从缓存读取：" + code)
                    console.log(`开始切分图片 ${sku}`)
                    result = await splitImage(fs.readFileSync(oriFile), true)
                    if(result == -2) {
                        console.log("图片已损坏，重新开始下载")
                    }
                }
                if(result != 1) {
                    const data = await downloadImage(imageUrl)
                    // if (lastSplitTask) {    // make sure last split task finished
                    //     await lastSplitTask
                    //     lastSplitTask = undefined
                    // }
                    if (data) {
                        console.log(`开始切分图片 ${sku}`)
                        result = await splitImage(data)
                        if(result == -2) {
                            console.log("下载的图片已损坏")
                        }
                    }
                        // lastSplitTask = splitImage(data)    // split and continue to download next
                }
            }
            // if (lastSplitTask) {
            //     await lastSplitTask
            //     lastSplitTask = undefined
            // }
            console.log(`开始生成条码 ${file}`)
            for (let order of orderMap.values()) {
                generateBarcode(order.code, [order.orderNumber, `Total: ${order.piece} pcs`, currentDate.toLocaleString()], path.join(outputFolderPath, `条码${order.code}.png`))
            }
        }
    }
    if(orderList.length == 0) {
        console.log("没有找到xlsx文件")
        return
    }
    const rowData = [
        "货品名称", "货品英文名称", "货品编号", "分类编号", "分类名称", "别名", "品牌", "单位", "辅助单位1", "转换率1", "辅助单位2", "转换率2", "辅助单位3", "转换率3", "常用单位", "规格", "条码", "长(cm)", "宽(cm)", "高(cm)", "体积", "体积单位", "重量单位", "重量", "固定成本价", "货品类型", "批次管理", "序列号管理", "生产物料", "定制生产", "提货卡券", "有偿服务", "需上门安装", "管理保质期", "保质期", "保质期单位", "货品标记", "规格标记", "备注", "货品说明", "在库生产", "序号", "规格备注", "自定义字段1", "自定义字段2", "规格编号", "颜色", "尺码", "成分", "文本", "测试日期01", "主条码", "默认供应商", "货主"
    ];


    for (let i = 0; i < orderList.length; i++) {
        const row = orderList[i]
        row[2] = `QHSTJ${year.toString().substring(2)}${month}${day}-${i + 8001}`
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


    const orderListFolder = "订单列表"
    if (!fs.existsSync(orderListFolder)) {
        fs.mkdirSync(orderListFolder);
    }
    const xlsFile = `${year}-${month}-${day}.xlsx`
    try {
        xlsx.writeFile(workbook, path.join(orderListFolder, xlsFile));
        console.log(`列表已生成 ${xlsFile}`)
    } catch (e) {
        console.log(`列表${xlsFile}无法覆盖，文件正在使用中`)
    }

}



function generateBarcode(code, messages, filename) {
    const barcodeOptions = {
        bcid: 'code128', // Barcode type
        text: code, // Text to encode
        scale: 4, // Barcode scaling factor
        height: 20, // Barcode height, in pixels
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

                    canvas.composite(barcodeImage, x, y);

                    // Add the number "15" at the bottom of the canvas

                    canvas.print(font, 100, 500, messages[0]);
                    canvas.print(font, 100, 580, messages[1]);
                    canvas.print(font2, 100, 660, messages[2]);

                    // Save the final image as a PNG file
                    canvas.write(filename, (saveErr) => {
                        if (saveErr) console.log("条码生成出错: " + code);
                    });
                });
            })
            .catch(barcodeImageErr => {
                console.error(barcodeImageErr);
            });
    });



}

process.on('uncaughtException', UncaughtExceptionHandler);

function UncaughtExceptionHandler(err) {
    console.log("err: ", err);
    console.log("Stack trace: ", err.stack);
}

main()