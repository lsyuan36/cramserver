const XLSX = require('xlsx');
const fs = require('fs');
const crypto = require('crypto');

// 生成隨機字符串
function generateRandomString(length = 3) {
    return crypto.randomBytes(length).toString('hex');
}

// 讀取 .xlsx 文件並將其轉換為 HTML
function convertXlsxSheetToHtml(filePath, sheetName) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_html(sheet);
}

// 獲取所有工作表的名稱
function getSheetNames(filePath) {
    const workbook = XLSX.readFile(filePath);
    return workbook.SheetNames;
}

// 生成映射表
function generateIdToSheetNameMapping(filePath) {
    const sheetNames = getSheetNames(filePath);
    const mapping = {};
    sheetNames.forEach(name => {
        const randomId = generateRandomString();
        mapping[randomId] = name;
    });
    return mapping;
}

// 加載或生成映射表
function loadOrCreateMapping(filePath, mappingFilePath) {
    if (fs.existsSync(mappingFilePath)) {
        return JSON.parse(fs.readFileSync(mappingFilePath));
    } else {
        const mapping = generateIdToSheetNameMapping(filePath);
        fs.writeFileSync(mappingFilePath, JSON.stringify(mapping, null, 2));
        return mapping;
    }
}

module.exports = {
    generateRandomString,
    convertXlsxSheetToHtml,
    getSheetNames,
    generateIdToSheetNameMapping,
    loadOrCreateMapping
};
