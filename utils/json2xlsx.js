// 导入json
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const Json2Xlsx = (jsonFilePath) => {
    const dirPath = path.dirname(jsonFilePath);
    if (!fs.existsSync(jsonFilePath)) {
        console.error("文件不存在，请提供正确的路径");
        process.exit(1); // 退出程序
    }
    const jsonData = require(path.resolve(jsonFilePath));
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Translation');
    worksheet.columns = [
        { header: '路径', key: 'keyPath', width: 80 },
        { header: '原始文案', key: 'originalText', width: 100 },
        { header: '翻译文案', key: 'translatedText', width: 50 }
    ];

    function parseJsonToExcel(json, parentKey = '') {
        for (const key in json) {
            if (typeof json[key] === 'object' && !Array.isArray(json[key])) {
                parseJsonToExcel(json[key], `${parentKey}${key}.`);
            } else if (Array.isArray(json[key])) {
                json[key].forEach((item, index) => {
                    if (typeof item === 'object') {
                        parseJsonToExcel(item, `${parentKey}${key}[${index}].`);
                    } else {
                        worksheet.addRow({
                            keyPath: `${parentKey}${key}[${index}]`,
                            originalText: item
                        });
                    }
                });
            } else {
                worksheet.addRow({
                    keyPath: `${parentKey}${key}`,
                    originalText: json[key]
                });
            }
        }
    }
    parseJsonToExcel(jsonData);
    const savePath = path.resolve(dirPath, 'translation_output.xlsx');
    workbook.xlsx.writeFile(savePath)
        .then(() => {
            console.log('Excel file created successfully! File path: ', savePath);
        })
        .catch((err) => {
            console.error('Error creating Excel file:', err);
        });
}

module.exports = Json2Xlsx;