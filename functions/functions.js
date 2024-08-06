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
function generateIdToSheetNameMapping(filePath, existingMapping = {}) {
    const sheetNames = getSheetNames(filePath);
    const mapping = { ...existingMapping };

    sheetNames.forEach(name => {
        if (!Object.values(mapping).includes(name)) {
            const randomId = generateRandomString();
            mapping[randomId] = name;
        }
    });

    return mapping;
}

// 加載或生成映射表
function loadOrCreateMapping(filePath, mappingFilePath) {
    let mapping = {};
    if (fs.existsSync(mappingFilePath)) {
        mapping = JSON.parse(fs.readFileSync(mappingFilePath));
    }
    const updatedMapping = generateIdToSheetNameMapping(filePath, mapping);
    fs.writeFileSync(mappingFilePath, JSON.stringify(updatedMapping, null, 2));
    return updatedMapping;
}

module.exports = {
    generateRandomString,
    convertXlsxSheetToHtml,
    getSheetNames,
    generateIdToSheetNameMapping,
    loadOrCreateMapping
};
