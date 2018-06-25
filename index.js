const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');
const config = require('./config.json');
const parser = require('./lib/process');
const glob = require('glob');

glob(config.xlsx.src, function (err, files) {
    if (err) {
        console.error("exportJson error:", err);
        throw err;
    }
    files.forEach(item => {
        toJson(path.join(__dirname, item), path.join(__dirname, config.xlsx.outPath));
    });
});

/**
 * 文件输出
 * @param {*} parsedWorkbook 文件数据
 * @param {*} outPath 输出路径
 */
function serializeWorkbook(parsedWorkbook, outPath) {

    for (let name in parsedWorkbook) {
        //json的名称
        let sheet = parsedWorkbook[name];
        //输出路径
        let outPathJson = path.resolve(outPath, name + ".json");
        //进行文本处理（压缩）
        let resultJson = JSON.stringify(sheet, null, config.json.compress ? 0 : 2); //, null, 2
        //写文件
        fs.writeFile(outPathJson, resultJson, err => {
            if (err) {
                console.error("error：", err);
                throw err;
            }
            console.log('exported sucessful -->  ', path.basename(outPathJson));
        });
    }
}

/**
 * 解析excel
 * @param {*} src excel path
 * @param {*} outPath  output json path
 */
function toJson(srcPath, outPath) {
    if (!fs.existsSync(outPath)) {
        fs.mkdirSync(outPath);
    }
    //解析后的原数据
    let workbook = xlsx.parse(srcPath);
    //处理后的数据
    let processData = parser.process(workbook);
    serializeWorkbook(processData, outPath);
}