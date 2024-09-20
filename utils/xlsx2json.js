const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');


const Excel2Json = (filePath) => {
    const dirPath = path.dirname(filePath);
    let newJsonData = {};

    function setValueByKeyPath(obj, keyPath, value) {
        const keys = keyPath.split('.');
        keys.reduce((acc, key, index) => {
            // 处理数组的情况，如 key[0] 这种格式
            const arrayMatch = key.match(/(\w+)\[(\d+)\]/);
            if (arrayMatch) {
                const arrayKey = arrayMatch[1];
                const arrayIndex = parseInt(arrayMatch[2], 10);
                acc[arrayKey] = acc[arrayKey] || [];
                acc[arrayKey][arrayIndex] = acc[arrayKey][arrayIndex] || {};
                if (index === keys.length - 1) {
                    acc[arrayKey][arrayIndex] = value;
                }
                return acc[arrayKey][arrayIndex];
            } else {
                if (index === keys.length - 1) {
                    acc[key] = value;  // 最后一级，赋值
                } else {
                    acc[key] = acc[key] || {};
                }
                return acc[key];
            }
        }, obj);
    }

    // 读取 Excel 文件并重建 JSON
    async function excelToJson() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        const worksheet = workbook.getWorksheet(1);  // 假设表格在第一个工作表

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {  // 跳过表头
                const keyPath = row.getCell(1).value;
                const translatedText = row.getCell(3).value;  // 第三列是翻译后的文本
                if (keyPath) {
                    setValueByKeyPath(newJsonData, keyPath, translatedText ? translatedText.toString() : "--");
                }
            }
        });
        const savePath = path.resolve(dirPath, 'en.json');

        fs.writeFileSync(savePath, JSON.stringify(newJsonData, null, 2), 'utf-8');
        console.log('Translated JSON file created successfully! \n' + savePath);
    }

    excelToJson().catch(console.error);

}

module.exports = Excel2Json;        