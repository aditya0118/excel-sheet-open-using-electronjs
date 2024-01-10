const express = require('express');
const app = express();
const bodyparser = require('body-parser');
const fs = require('fs');
const readXlsxFile = require('read-excel-file/node');
const mysql = require('mysql');
const multer = require('multer');
const path = require('path');

const __basedir = path.resolve(__dirname);

app.use(express.static(path.join(__dirname, 'public')));

app.use(bodyparser.json());
app.use(bodyparser.urlencoded({
    extended: true
}));

const db = mysql.createConnection({
    host: "127.0.0.1",
    user: "root",
    password: "",
    database: "demo"
});

db.connect(function (err) {
    if (err) {
        return console.error('error: ' + err.message);
    }
    console.log('Connected to the MySQL server.');
});

// Create 'uploads' directory if it does not exist
fs.mkdirSync(path.join(__basedir, 'uploads'), { recursive: true });

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, path.join(__basedir, 'uploads'));
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + "-" + file.originalname);
    }
});

const upload = multer({ storage: storage });

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '/index.html'));
});

app.post('/uploadfile', upload.single("uploadfile"), (req, res) => {
    const uploadPath = path.join(__basedir, 'uploads', req.file.filename);
    importExcelData2MySQL(uploadPath, res);
});

function importExcelData2MySQL(filePath, res) {
    const headerMapping = {
        'Cust_Id': 'id',
        'Cust_Add': 'address',
        'Cust_Age': 'age',
        'Cust_Name': 'name'
    };

    readXlsxFile(filePath).then((rows) => {
        const headers = rows.shift(); // Extract headers from the first row
        const mappedHeaders = headers.map(header => headerMapping[header] || header);

        const mappedRows = rows.map(row => {
            const mappedRow = {};
            row.forEach((value, index) => {
                mappedRow[mappedHeaders[index]] = value;
            });
            return mappedRow;
        });

        db.query('INSERT INTO customer (id, address, age, name) VALUES ?', [mappedRows.map(Object.values)], (error, response) => {
            if (error) {
                console.error('Database error:', error);
                res.status(500).send('Internal Server Error');
            } else {
                console.log('Database response:', response);
                res.send('File uploaded successfully.');
            }
        });
    });
}

let server = app.listen(8080, function () {
    let host = server.address().address;
    let port = server.address().port;
    console.log("App listening at http://%s:%s", host, port);
});