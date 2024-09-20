const express = require('express');
const path = require('path');
const fs = require('fs');
const puppeteer = require('puppeteer');
const archiver = require('archiver');
const multer = require('multer');
const { convertXlsxSheetToHtml, loadOrCreateMapping, generateIdToSheetNameMapping, getSheetNames } = require('./functions/functions');
const { processFile , salary } = require('./functions/python_call');

const app = express();
const port = 5000;
const current_month = '2024年8月份';

// 设置静态文件夹
app.use(express.static(path.join(__dirname, 'public')));

// 中间件：解析 URL 编码的请求体
app.use(express.urlencoded({ extended: true }));

const file1Path = path.join(__dirname, 'output/老師個別表.xlsx');
const file2Path = path.join(__dirname, 'output/學生個別表.xlsx');
const mappingFile1Path = path.join(__dirname, 'mapping/teacher.json');
const mappingFile2Path = path.join(__dirname, 'mapping/student.json');
let idToSheetName1;
let idToSheetName2;

try {
    idToSheetName1 = loadOrCreateMapping(file1Path, mappingFile1Path);
    idToSheetName2 = loadOrCreateMapping(file2Path, mappingFile2Path);
}catch (e){
    console.log("ID Failed at initial")
}
// 创建目录（如果不存在）
function createDirectoryIfNotExists(directory) {
    if (!fs.existsSync(directory)) {
        fs.mkdirSync(directory, { recursive: true });
    }
}

createDirectoryIfNotExists(path.join(__dirname, 'output/老師'));
createDirectoryIfNotExists(path.join(__dirname, 'output/學生'));
createDirectoryIfNotExists(path.join(__dirname, 'input'));


// 配置multer用于文件上传
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, path.join(__dirname, 'input'));
    },
    filename: function (req, file, cb) {
        cb(null, '上課記錄 (回覆).xlsx');
    }
});

const upload = multer({ storage: storage });

const studentUploadStorage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, path.join(__dirname, 'output'));
    },
    filename: function (req, file, cb) {
        cb(null, '學生個別表.xlsx');
    }
});
const teacherUploadStorage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, path.join(__dirname, 'output'));
    },
    filename: function (req, file, cb) {
        cb(null, '老師個別表.xlsx');
    }
});
const teacherbasedsalaryUploadStorage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, path.join(__dirname, 'input'));
    },
    filename: function (req, file, cb) {
        cb(null, '基底薪資.xlsx');
    }
});



const studentUpload = multer({ storage: studentUploadStorage });
const teacherUpload = multer({ storage: teacherUploadStorage });
const teacherbasedsalary = multer({storage: teacherbasedsalaryUploadStorage });

// 主页路由处理
app.get('/controller', (req, res) => {
    let html = `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>上課清單</title>
            <link rel="stylesheet" href="/styles.css">
        </head>
        <body>
            <div class="container">
                <h1>上課清單</h1>
                <table>
                    <td>
                        <a href = "https://docs.google.com/spreadsheets/d/1q3JYAD0L0Kj9c47PbNPI0wvC8bbtWz1PAGYChWXg0Ps/edit?usp=sharing">
                            <button>前往Google 試算表</button></a>
                        <form action="/upload-and-process" method="post" enctype="multipart/form-data">
                            
                            <div>
                                <label for="file">上傳新文件：</label>
                                <input type="file" name="file" id="file" required>
                            </div>
                            <div>
                                <label for="date">輸入日期 (yyyy/mm)：</label>
                                <input type="text" name="date" id="date" placeholder="2024/08" required>
                            </div>
                            <button type="submit">上傳並處理文件</button>
                        </form>
                    </td>
                    <td>
                        <form action="/upload-based-salary" method="post" enctype="multipart/form-data">
                            <div>
                                <button onclick="window.location.href='/download/teacher-based-salary'">下載目前薪資基底表.xlsx</button><br>
                                <label for="file">上傳薪資基底表：</label>
                                <input type="file" name="basedSalaryFile" id="basedSalaryFile" required>
                            </div>
                            <button type="submit">上傳並處理文件</button>
                        </form>
                    </td>
                </table>

                <h2>老師個別表</h2>
                <table>
                    <tbody>
                        <tr>
    `;

    Object.keys(idToSheetName1).forEach((id, index) => {
        html += `<td><a href="/view1/${id}">${idToSheetName1[id]}</a></td>`;
        if ((index + 1) % 10 === 0) {
            html += `</tr><tr>`;
        }
    });

    html += `
                        </tr>
                    </tbody>
                </table>
                <table>
                    <td>
                        <button onclick="window.location.href='/generate-download-teachers-zip'">下載所有老師的PDF壓縮包</button>
                    </td>
                    <td>
                        <button onclick="window.location.href='/download/teacher-excel'">下載老師個別表.xlsx</button>
                    </td>
                    <td>
                        <button onclick="window.location.href='/download/teacher-salary'">下載老師薪水總表.xlsx</button>
                    </td>
                    <td>
                        <form action="/upload-teacher-excel" method="post" enctype="multipart/form-data">
                            <div>
                                <label for="teacherFile">上傳修改後的老師個別表：</label>
                                <input type="file" name="teacherFile" id="teacherFile" required>
                            </div>
                            <button type="submit">上傳並重新載入老師生資料</button>
                        </form>   
                    </td>
                </table>    

                <h2>學生個別表</h2>
                <table>
                    <tbody>
                        <tr>
    `;

    Object.keys(idToSheetName2).forEach((id, index) => {
        html += `<td><a href="/view2/${id}">${idToSheetName2[id]}</a></td>`;
        if ((index + 1) % 10 === 0) {
            html += `</tr><tr>`;
        }
    });

    html += `
                        </tr>
                    </tbody>
                </table>
                <table>
                    <td>
                        <button onclick="window.location.href='/generate-download-students-zip'">下載所有學生的PDF壓縮包</button>                    
                    </td>
                    <td>
                         <button onclick="window.location.href='/download/student-excel'">下載學生個別表.xlsx</button>
                    </td>
                    <td>
                        <form action="/upload-student-excel" method="post" enctype="multipart/form-data">
                            <div>
                                <label for="studentFile">上傳修改後的學生個別表：</label>
                                <input type="file" name="studentFile" id="studentFile" required>
                            </div>
                            <button type="submit">上傳並重新載入學生資料</button>
                        </form>
                    </td>                                        
                </table>
            </div>
        </body>
        </html>
    `;

    res.send(html);
});

// 处理文件上传和处理
app.post('/upload-and-process', upload.single('file'), (req, res) => {
    const date = req.body.date;

    processFile(date).then(r => {
        console.log(r);
        idToSheetName1 = loadOrCreateMapping(file1Path, mappingFile1Path, true);
        idToSheetName2 = loadOrCreateMapping(file2Path, mappingFile2Path, true);
        res.redirect('/controller');
    }).catch(err => {
        console.error(err);
        res.status(500).send('文件处理失败');
    });
});

// 处理上传修改后的学生个别表并重新载入
app.post('/upload-student-excel', studentUpload.single('studentFile'), (req, res) => {
    const newFilePath = path.join(__dirname, 'output/學生個別表.xlsx');
    if (fs.existsSync(newFilePath)) {
        idToSheetName2 = generateIdToSheetNameMapping(newFilePath);
        fs.writeFileSync(mappingFile2Path, JSON.stringify(idToSheetName2));
    }
    res.redirect('/controller');
});

// 处理上传修改后的老師个别表并重新载入
app.post('/upload-teacher-excel', teacherUpload.single('teacherFile'), (req, res) => {
    const newFilePath1 = path.join(__dirname, 'output/老師個別表.xlsx');
    if (fs.existsSync(newFilePath1)) {
        idToSheetName1 = generateIdToSheetNameMapping(newFilePath1);
        fs.writeFileSync(mappingFile1Path, JSON.stringify(idToSheetName1));
    }
    res.redirect('/controller');
});

// 处理上传修改后的基底薪資并重新载入
app.post('/upload-based-salary', teacherbasedsalary.single('basedSalaryFile'), (req, res) => {
    // const newFilePath2 = path.join(__dirname, 'input/基底薪資.xlsx');
    salary().then(r => console.log(r));
    res.redirect('/controller');
});

// 显示老师个别表的工作表内容的路由处理
app.get('/view1/:id', (req, res) => {
    const id = req.params.id;

    if (idToSheetName1[id]) {
        const sheetName = idToSheetName1[id];
        let htmlContent = convertXlsxSheetToHtml(file1Path, sheetName);
        htmlContent = htmlContent.replace(/<table>/, '<table class="table">'); // 添加 CSS 类
        res.send(`
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>${sheetName}</title>
                <link rel="stylesheet" href="/styles.css">
            </head>
            <body>
                <div class="container">
                    <h1>${current_month} ${sheetName} 薪水表</h1>
                    ${htmlContent}
                </div>
            </body>
            </html>
        `);
    } else {
        res.status(404).send('Sheet not found');
    }
});

// 显示学生个别表的工作表内容的路由处理
app.get('/view2/:id', (req, res) => {
    const id = req.params.id;

    if (idToSheetName2[id]) {
        const sheetName = idToSheetName2[id];
        let htmlContent = convertXlsxSheetToHtml(file2Path, sheetName);
        htmlContent = htmlContent.replace(/<table>/, '<table class="table">'); // 添加 CSS 类
        res.send(`
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>${sheetName}</title>
                <link rel="stylesheet" href="/styles.css">
            </head>
            <body>
                <div class="container">
                    <h1>${current_month} ${sheetName} 上課記錄</h1>
                    ${htmlContent}
                </div>
            </body>
            </html>
        `);
    } else {
        res.status(404).send('Sheet not found');
    }
});

// 生成老师个别表的所有页面的 PDF 并下载 ZIP
app.get('/generate-download-teachers-zip', async (req, res) => {
    const browser = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox']});
    const page = await browser.newPage();

    for (const id in idToSheetName1) {
        const sheetName = idToSheetName1[id];
        const htmlContent = convertXlsxSheetToHtml(file1Path, sheetName);
        const content = `
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>${sheetName}</title>
                <link rel="stylesheet" href="/styles.css">
                <style>
                    h1{
                        text-align: center;
                    }
                    table, th, td {
                        border: 1px solid black;
                        border-collapse: collapse;
                        text-align: center;
                    }
                    table{
                        margin-left:auto; 
                        margin-right:auto;
                    }
                    body {
                        font-size: 12pt;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>${sheetName}</h1>
                    ${htmlContent.replace(/<table>/, '<table class="table">')}
                </div>
            </body>
            </html>
        `;
        await page.setContent(content);
        await page.pdf({ path: `output/老師/${current_month}_${sheetName}_薪水表.pdf`, format: 'A4', landscape: true });
    }

    await browser.close();

    const output = fs.createWriteStream(path.join(__dirname, 'output/老師.zip'));
    const archive = archiver('zip', {
        zlib: { level: 9 } // 设置压缩级别
    });

    output.on('close', () => {
        res.download(path.join(__dirname, 'output/老師.zip'));
    });

    archive.on('error', (err) => {
        throw err;
    });

    archive.pipe(output);

    fs.readdir(path.join(__dirname, 'output/老師'), (err, files) => {
        if (err) {
            console.error('Error reading directory:', err);
            res.status(500).send('Server error');
            return;
        }

        files.forEach((file) => {
            archive.file(path.join(__dirname, 'output/老師', file), { name: file });
        });

        archive.finalize();
    });
});

// 生成学生个别表的所有页面的 PDF 并下载 ZIP
app.get('/generate-download-students-zip', async (req, res) => {
    const browser = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox']});
    const page = await browser.newPage();

    for (const id in idToSheetName2) {
        const sheetName = idToSheetName2[id];
        const htmlContent = convertXlsxSheetToHtml(file2Path, sheetName);
        const content = `
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>${sheetName}</title>
                <link rel="stylesheet" href="/styles.css">
                <style>
                    h1{
                        text-align: center;
                    }
                    table, th, td {
                        border: 1px solid black;
                        border-collapse: collapse;
                        text-align: center;
                    }
                    table{
                        margin-left:auto; 
                        margin-right:auto;
                    }
                    body {
                        font-size: 12pt;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>${current_month} ${sheetName} 上課記錄</h1>
                    ${htmlContent.replace(/<table>/, '<table class="table">')}
                </div>
            </body>
            </html>
        `;
        await page.setContent(content);
        await page.pdf({ path: `output/學生/${current_month}_${sheetName}.pdf`, format: 'A4', landscape: true });
    }

    await browser.close();

    const output = fs.createWriteStream(path.join(__dirname, 'output/學生.zip'));
    const archive = archiver('zip', {
        zlib: { level: 9 } // 设置压缩级别
    });

    output.on('close', () => {
        res.download(path.join(__dirname, 'output/學生.zip'));
    });

    archive.on('error', (err) => {
        throw err;
    });

    archive.pipe(output);

    fs.readdir(path.join(__dirname, 'output/學生'), (err, files) => {
        if (err) {
            console.error('Error reading directory:', err);
            res.status(500).send('Server error');
            return;
        }

        files.forEach((file) => {
            archive.file(path.join(__dirname, 'output/學生', file), { name: file });
        });

        archive.finalize();
    });
});

// 下载老师个别表 Excel
app.get('/download/teacher-excel', (req, res) => {
    const file = path.join(__dirname, 'output/老師個別表.xlsx');
    res.download(file, err => {
        if (err) {
            console.error('Error downloading file:', err);
            res.status(500).send('Server error');
        }
    });
});

// 下载学生个别表 Excel
app.get('/download/student-excel', (req, res) => {
    const file = path.join(__dirname, 'output/學生個別表.xlsx');
    res.download(file, err => {
        if (err) {
            console.error('Error downloading file:', err);
            res.status(500).send('Server error');
        }
    });
});

// 下载老師薪水總表
app.get('/download/teacher-salary', (req, res) => {
    const file = path.join(__dirname, 'output/薪水總表.xlsx');
    res.download(file, err => {
        if (err) {
            console.error('Error downloading file:', err);
            res.status(500).send('Server error');
        }
    });
});

// 下载基底薪資
app.get('/download/teacher-based-salary', (req, res) => {
    const file = path.join(__dirname, 'input/基底薪資.xlsx');
    res.download(file, err => {
        if (err) {
            console.error('Error downloading file:', err);
            res.status(500).send('Server error');
        }
    });
});

app.get('/', (req, res) => {
     res.send('Server ON');
});

app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}/controller`);
});
