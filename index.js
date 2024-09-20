const inquirer = require('inquirer');
const Json2Xlsx = require('./utils/json2xlsx');
const Excel2Json = require('./utils/xlsx2json');
inquirer
    .prompt([
        {
            type: 'list',
            name: 'type',
            choices: ['JSON转Excel', 'Excel转Json'],
            message: '功能选择:'
        },
        {
            type: "input",
            name: "filePath",
            message: '文件路径:（转换文件将生成在当前文件夹下）'
        }

    ])
    .then(answers => {
        const filePath = answers.filePath;
        if (answers.type === 'JSON转Excel') {
            Json2Xlsx(filePath);
        } else {
            Excel2Json(filePath);
        }
    })
    .catch(error => {
        console.error('发生错误:', error);
    });
