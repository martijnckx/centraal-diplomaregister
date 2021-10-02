const express = require('express')
const xlsx = require('node-xlsx')
const multer = require('multer')
const https = require('https')
const path = require('path')
const fs = require('fs')
const app = express()
const port = 3000

// source: https://github.com/oleg-koval/excel-date-to-js/blob/master/index.js
const excelDateToJS = function(excelDate) {
    const SECONDS_IN_DAY = 24 * 60 * 60;
    const MISSING_LEAP_YEAR_DAY = SECONDS_IN_DAY * 1000;
    const MAGIC_NUMBER_OF_DAYS = (25567 + 2);

    const delta = excelDate - MAGIC_NUMBER_OF_DAYS;
    const parsed = delta * MISSING_LEAP_YEAR_DAY;
    return new Date(parsed);
}

const getCsrfTokens = function() {
    const options = {
        hostname: 'csm-examen.be',
        port: 443,
        path: '/cdr',
        method: 'GET',
    }
    return new Promise(resolve => {
        const req = https.request(options, res => {
            let data = '';

            res.on('data', d => {
                data += d
            })

            res.on('end', () => {
                const tokenRegex = new RegExp(/name="_token" type="hidden" value="([\w\d]{40})">/g)
                const csrfToken = tokenRegex.exec(data)[1]
                resolve({
                    csrf: csrfToken,
                    cookies: res.headers['set-cookie'].map(h => h.split(';')[0]).join('; '),
                })
            });
        })

        req.on('error', error => {
            console.error(error)
        })

        req.end()
    })
}

const getDegreeInfo = async function(person, tokens) {
    const data = `_token=${tokens.csrf}&diplomaNumber=&lastName=${person[0]}&dateOfBirth=${person[1]}`;
    const options = {
        hostname: 'csm-examen.be',
        port: 443,
        path: '/cdr',
        headers: {
            'Cookie': tokens.cookies,
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': Buffer.byteLength(data),
        },
        method: 'POST',
    }

    return new Promise(resolve => {
        // Set up the request
        var req = https.request(options, res => {
            let data = '';

            res.on('data', d => {
                data += d
            })

            res.on('end', () => {
                const degreeRegex = new RegExp(/<dt>Type<\/dt>\s*<dd>(.*)<\/dd>\s*<dt>Geldig tot<\/dt>\s*<dd>\s*(\d{1,2}-\d{1,2}-\d{4})\s*<\/dd>/g)
                const degreeInfo = degreeRegex.exec(data);
                let result = ['/', '/'];
                if (degreeInfo) result = [degreeInfo[1], degreeInfo[2]]
                resolve(result)
            });
        });

        req.on('error', error => {
            console.error(error)
        })

        // post the data
        req.write(data);
        req.end();

    });

    // _token: SNkubzlhPMhD7b2nNXavk2jGJKw3ef5dyCKQGKTu
    // diplomaNumber: 
    // lastName: Patteeuw
    // dateOfBirth: 05-04-1976
}

const doLookups = async function(filepath) {
    // Format correctly and get the desired columns
    let sheet_original = xlsx.parse(filepath)[0].data
    sheet_original = sheet_original.map(a => {
        if (a[a.length - 1] !== 'Geboortedag') {
            const dateobj = excelDateToJS(a[a.length - 1]);
            a[a.length - 1] = dateobj;
        }
        return a;
    })
    let sheet = xlsx.parse(filepath)[0].data;
    sheet.shift();
    sheet = sheet.map(a => {
        const dateobj = excelDateToJS(a[a.length - 1]);
        a[a.length - 1] = `${dateobj.getDate()}-${dateobj.getMonth()+1}-${dateobj.getFullYear()}`;
        return [a[3], a[5]];
    });
    sheet = sheet.filter(x => x && x[0] && x[1])
        // Get CSRF token
    const tokens = await getCsrfTokens();
    let degrees = [];
    for (let person of sheet) {
        let degree = await getDegreeInfo(person, tokens);
        degrees.push(degree)
    }

    sheet_original[0][6] = 'Titel diploma';
    sheet_original[0][7] = 'Geldigheid diploma';

    i = 0;
    while (i < sheet_original.length - 1 && i < degrees.length) {
        sheet_original[i + 1][6] = degrees[i][0];
        sheet_original[i + 1][7] = degrees[i][1];
        i++;
    }
    const newSheetData = xlsx.build([{ name: "withDegrees", data: sheet_original, type: 'base64' }]); // Returns a buffer
    return newSheetData;
}

const makeid = function(length) {
    var result = '';
    var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var charactersLength = characters.length;
    for (var i = 0; i < length; i++) {
        result += characters.charAt(Math.floor(Math.random() *
            charactersLength));
    }
    return result;
}

const onlySheets = function(req, file, cb) {
    // Accept images only
    if (!file.originalname.match(/\.(xlsx)$/) || file.mimetype !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        const errorMsg = 'Only Excel sheets are allowed (.xlsx)!';
        req.fileValidationError = errorMsg;
        return cb(new Error(errorMsg), false);
    }
    cb(null, true);
};

// Serve everything from 'public' folder as-is
app.use(express.static(path.join(__dirname, 'public')))

const storage = multer.diskStorage({
    destination: function(req, file, cb) {
        cb(null, 'uploads/');
    },

    // By default, multer removes file extensions so let's add them back
    filename: function(req, file, cb) {
        cb(null, `${Date.now()}-${makeid(4)}${path.extname(file.originalname)}`);
    }
});

// Show homepage
app.get('/', (req, res, next) => {
    res.sendFile(path.join(__dirname, 'index.html'))
})

// Receive upload and start lookups
app.post('/lookup', (req, res, next) => {
    let upload = multer({ storage: storage, fileFilter: onlySheets }).single('sheet');

    upload(req, res, async function(err) {
        // req.file contains information of uploaded file
        // req.body contains information of text fields, if there were any

        if (req.fileValidationError) {
            return res.send(req.fileValidationError);
        } else if (!req.file) {
            return res.send('Please select an excel sheet to upload');
        } else if (err instanceof multer.MulterError) {
            return res.send(err);
        } else if (err) {
            return res.send(err);
        }

        const lookups = await doLookups(req.file.path);
        fs.unlink(req.file.path, (err) => {
            if (err) throw err;
        })

        res.writeHead(200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        res.end(Buffer.from(lookups, 'base64'));
    });
})

app.get('*', function(req, res) {
    res.type('txt').status(404).send(`404 - Not Found`);
});


app.listen(port, () => {
    console.log(`App listening at http://localhost:${port}`)
})