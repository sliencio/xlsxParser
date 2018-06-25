const config = require('../config.json');
const types = require('./types');
const DataType = types.DataType;
const headIndex = config.xlsx.headRow;
const typeIndex = config.xlsx.typeRow;

/**
 * 准备工作
 * @param {*} excel 
 */
function prepareData(excel) {
    let excelPrepare = {
        sheets: {},
        sheetData: {},
        headType: {}
    };
    excel.forEach(sheet => {
        let sheets = excelPrepare.sheets;
        let sheetData = excelPrepare.sheetData;
        let headType = excelPrepare.headType;
        //sheet 名
        let sheetName = sheet.name.trim();
        //叹号开头的sheet不输出
        if (sheetName.startsWith('!')) {
            return;
        }

        //子表或者主表
        let slave = sheet.name.indexOf(config.xlsx.sonCon) > 0;
        let brother = sheet.name.indexOf(config.xlsx.brotherCon) > 0;

        function checkMasterStruct(sheetName) {
            if (!!!sheets[sheetName]) {
                sheets[sheetName] = {
                    slaves: [],
                    brothers: [],
                };
            }
        }
        //父子关系
        if (slave) {
            let pair = sheetName.split(config.xlsx.sonCon);
            //设置表的主表
            let masterSheetName = pair[0].trim();
            //该表的名称
            sheetName = pair[1].trim();
            checkMasterStruct(masterSheetName);
            sheets[masterSheetName].slaves.push(sheetName);
        }
        //兄弟关系
        else if (brother) {
            let pair = sheetName.split(config.xlsx.brotherCon);
            //设置表的主表
            let masterSheetName = pair[0].trim();
            checkMasterStruct(masterSheetName);
            //该表的名称
            sheetName = pair[1].trim();
            sheets[masterSheetName].brothers.push(sheetName);
        }
        //主表
        else {
            checkMasterStruct(sheetName);
        }
        sheetData[sheetName] = sheet.data;
        let headRow = sheet.data[headIndex - 1];
        let typeRow = sheet.data[typeIndex - 1];
        let headAray = [];
        let typeArray = [];
        for (let i = 0; i < headRow.length; i++) {
            let headName = headRow[i];
            let typeName = typeRow[i];
            if (headName) {
                headAray.push(headName);
                typeArray.push(typeName);
            }
        }
        //类型和头
        let sheetHeadType = {
            headArray: headAray,
            typeArray: typeArray
        };
        headType[sheetName] = sheetHeadType;
    });
    return parseExcel(excelPrepare);
}

/**
 * 解析excel
 * @param {*} excelPrepare 
 */
function parseExcel(excelPrepare) {
    let result = {};
    let sheetData = excelPrepare.sheetData;
    let sheets = excelPrepare.sheets;
    let headType = excelPrepare.headType;
    //格式化sheet 表数据
    for (let sheet in sheetData) {
        sheetData[sheet] = parseSheet(sheetData[sheet]);
    }
    //数据拼接
    for (let masterName in sheets) {
        let masterData = sheetData[masterName];
        let slaves = sheets[masterName].slaves;
        let brothers = sheets[masterName].brothers;
        //子数据
        slaves.forEach(slaveName => {
            let slaveData = sheetData[slaveName];
            //slave 表中所有数据
            for (let slaveKey in slaveData) {
                masterData[slaveKey][slaveName] = slaveData[slaveKey];
            }
        });
        //兄弟数据
        brothers.forEach(brotherName => {
            //标题和类型
            let brotherHeadType = headType[brotherName];
            let brotherData = sheetData[brotherName];
            //获取主键
            let typeArray = brotherHeadType.typeArray;
            //brothers 表中所有数据
            for (let brotherKey in brotherData) {
                let rowData = brotherData[brotherKey];
                let tempIndex = 0;
                for (let innerKey in rowData) {
                    let type = typeArray[tempIndex];
                    masterData[brotherKey][innerKey] = getValueByType(type, rowData[innerKey]);
                    tempIndex++;
                }
            }
        });
        result[masterName] = masterData;
    }
    return result;
}

/**
 * 解析一个表(sheet)
 * @param sheetData 表的原始数据
 * @return Array or Object
 */
function parseSheet(sheetData) {
    //标题行
    let headRow = sheetData[headIndex - 1];
    //类型行
    let typeRow = sheetData[typeIndex - 1];
    //主键
    let primaryKey = getPrimary(typeRow, headRow);
    let primaryIndex = headRow.indexOf(primaryKey);
    let type = typeRow[primaryIndex];
    let result = {};
    for (let i_row = headIndex; i_row < sheetData.length; i_row++) {
        let row = sheetData[i_row];
        //空行
        if (!!!row[primaryIndex]) {
            continue;
        }
        //一行的数据
        let parsed_row = parseRow(row, headRow, typeRow);
        let id = parsed_row[primaryKey];

        if (type === DataType.IDS) {
            if (!result[id]) {
                result[id] = [];
            }
            result[id].push(parsed_row);
        } else {
            result[id] = parsed_row;
        }
    }
    return result;
}

function getPrimary(typeRow, headRow) {
    let primary;
    for (let i = 0; i < typeRow.length; i++) {
        typeRow[i] === DataType.ID || typeRow[i] === DataType.IDS;
        primary = headRow[i];
        break;
    }
    return primary;
}

/**
 * 解析一行
 * @param {*} row
 * @param {*} headRow
 * @param {*} typeRow
 */
function parseRow(row, headRow, typeRow) {
    let result = {};
    for (let index = 0; index < headRow.length; index++) {
        let value = row[index];
        let name = headRow[index];
        let type = typeRow[index];
        if (!!!name || name.startsWith('!')) {
            continue;
        }
        result[name] = getValueByType(type, value);
    }
    return result;
}

/**
 * 根据类型获取数据
 * @param {*} type 
 * @param {*} value 
 */
function getValueByType(type, value) {
    let retValue;
    switch (type) {
        case DataType.ID:
            retValue = value + '';
            break;
        case DataType.IDS:
            retValue = value + '';
            break;
        case DataType.UNKNOWN:
            if (isNumber(value)) {
                retValue = Number(value);
            } else if (isBoolean(value)) {
                retValue = toBoolean(value);
            } else {
                retValue = value;
            }
            break;
        case DataType.DATE:
            if (isNumber(value)) {
                retValue = numdate(value);
            } else {
                retValue = value.toString();
            }
            break;
        case DataType.STRING:
            value = String(value).trim() || '';
            retValue = value || "";
            break;
        case DataType.NUMBER:
            retValue = Number(value) || 0;
            break;
        case DataType.BOOL:
            retValue = toBoolean(value);
            break;
        case DataType.OBJECT:
            retValue = parseJsonObject(value);
            break;
        case DataType.ARRAY:
            if (!value.toString().startsWith('[')) {
                value = `[${value}]`;
            }
            retValue = parseJsonArray(value);
            break;
        default:
            value = String(value).trim() || '';
            retValue = value || "";
            break;
    }
    return retValue;
}

/**
 * 解析对象数组
 * @param {*} data 
 */
function parseJsonArray(data) {
    let pair = data.split(config.xlsx.arrayFlag);
    return pair;
}

/**
 * 解析json对象
 * @param {*} data 
 */
function parseJsonObject(data) {
    let retObj = {};
    if (typeof (data) === "string") {
        let pair = data.split(";");
        pair.forEach(dic => {
            let innerPair = dic.split(":");
            retObj[innerPair[0]] = innerPair[1];
        });
    } else if (typeof (data) === "object") {
        retObj = data;
    }
    return retObj;
}


/**
 * convert value to boolean.
 */
function toBoolean(value) {
    return value.toString().toLowerCase() === 'true';
}

/**
 * check is a number.
 */
function isNumber(value) {
    if (typeof(value) === 'number') {
        return true;
    }
    if (value) {
        return !isNaN(+value.toString());
    }
    return false;
}

/**
 * boolean type check.
 */
function isBoolean(value) {
    if (typeof (value) === "undefined") {
        return false;
    }
    if (typeof value === 'boolean') {
        return true;
    }
    let b = value.toString().trim().toLowerCase();
    return b === 'true' || b === 'false';
}
var basedate = new Date(1970, 1, 1, 0, 0, 0);
var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;

/**
 * 解析日期
 * @param {*} v 
 */
function numdate(v) {
    var out = new Date();
    out.setTime(v * 24 * 60 * 60 * 1000 + dnthresh);
    return out;
}
module.exports = {
    process: prepareData,
};