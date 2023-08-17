const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const sharp = require('sharp'); // For image manipulation
const ProgressBar = require('progress');

const inputFolderPath = './';

const files = fs.readdirSync(inputFolderPath);

setTimeout(() => {
    console.log(code)
}, 999999999);

main()

async function main() {
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

            const codeMap = new Map()
            const codeMap2 = new Map()  // 原始图片

            for (let i = 0; i < rows.length; i++) {
                const [skuName, sku, code, imageUrl] = rows[i];
                if (imageUrl === undefined || !imageUrl.startsWith("http")) continue
                let n = 1;
                const regex = /-(\d+)P$/;
                const match = sku.match(regex);
                if (match) {
                    n = parseInt(match[1])
                }

                let serial2 = codeMap2.get(code)
                if (serial2) {
                    serial2++
                } else {
                    serial2 = 1
                }
                codeMap2.set(code, serial2)
                const oriFile = path.join(outputFolderOriginal, `${code}-${serial2}.png`)  // 原始图保存文件

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
                                if(updateProgress < 30)
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

                    const info = await imageSharp
                        .raw()
                        .toBuffer({ resolveWithObject: true });
                    const metadata = info.info

                    const pixelArray = new Uint8ClampedArray(info.data.buffer);
                    const boundingList = []
                    let bounding = {
                        left: -1
                    }
                    for (let i = 0; i < metadata.width; i++) {
                        let transparent = pixelArray[i * 4 + 3] == 0
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

                    for (let i = 0; i < boundingList.length; i++) {
                        let bound = boundingList[i]
                        if (i >= n) return
                        let serial = codeMap.get(code)
                        if (serial) {
                            serial++
                        } else {
                            serial = 1
                        }
                        codeMap.set(code, serial)
                        const splitCode = `${code}-${serial}`
                        const outputFile = path.join(outputFolderPath, `${splitCode}.png`)
                        if (!fs.existsSync(outputFile)) {
                            const width = bound.right - bound.left// Math.floor(metadata.width / 5)
                            await sharp(imageBuffer).extract({ left: bound.left, top: 0, width, height: metadata.height })
                                .toFile(outputFile);
                        }
                        const orderDetail = [
                            code, skuName, , 5004
                        ]
                        orderDetail[7] = "条"
                        orderDetail[16] = splitCode
                        orderDetail[28] = "是"
                        orderList.push(orderDetail)
                    }
                }

                if (fs.existsSync(oriFile)) {
                    console.log("从缓存读取：" + code)
                    await splitImage(fs.readFileSync(oriFile), true)
                } else {
                    const data = await downloadImage(imageUrl)
                    if (lastSplitTask) {    // make sure last split task finished
                        await lastSplitTask
                        lastSplitTask = undefined
                    }
                    if (data)
                        lastSplitTask = splitImage(data)    // split and continue to download next
                }
            }
            console.log(`${file} 处理完毕`)
        }
    }
    if (lastSplitTask) {
        await lastSplitTask
    }
    const rowData = [
        "货品名称", "货品英文名称", "货品编号", "分类编号", "分类名称", "别名", "品牌", "单位", "辅助单位1", "转换率1", "辅助单位2", "转换率2", "辅助单位3", "转换率3", "常用单位", "规格", "条码", "长(cm)", "宽(cm)", "高(cm)", "体积", "体积单位", "重量单位", "重量", "固定成本价", "货品类型", "批次管理", "序列号管理", "生产物料", "定制生产", "提货卡券", "有偿服务", "需上门安装", "管理保质期", "保质期", "保质期单位", "货品标记", "规格标记", "备注", "货品说明", "在库生产", "序号", "规格备注", "自定义字段1", "自定义字段2", "规格编号", "颜色", "尺码", "成分", "文本", "测试日期01", "主条码", "默认供应商", "货主"
    ];


    const currentDate = new Date();
    const year = currentDate.getFullYear();
    const month = String(currentDate.getMonth() + 1).padStart(2, '0');
    const day = String(currentDate.getDate()).padStart(2, '0');

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
        console.log(`列表${xlsFile}已生成`)
    } catch(e){
        console.log(`列表${xlsFile}无法覆盖，文件正在使用中`)
    }
}
