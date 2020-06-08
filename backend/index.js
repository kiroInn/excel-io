const express = require('express')
const formidable = require('formidable');
const path = require('path');
const xlsx = require('node-xlsx');
const fs = require('fs');
const stream = require('stream');

const app = express();

app.use(function(req, res, next) {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
  next();
});

app.use('/static', express.static(__dirname +'/template'))

app.get('/index', (req, res) => {
  res.sendFile(path.join(__dirname + '/../frontend/dist/index.html'));
});

app.post('/api/upload', (req, res, next) => {
  const form = formidable({ multiples: true });
  form.parse(req, (err, fields, files) => {
    if (err) {
      next(err);
      return;
    }
    const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(files.file.path))
    // console.log(workSheetsFromBuffer)
    console.log('sheets', workSheetsFromBuffer.Sheets['Reconciliation'])
    // builder-106700();
    res.json({ fields, files });
  });
});

app.get('/api/download', (request, response) => {
  const fileData = 'SGVsbG8sIFdvcmxkIQ=='
  const fileName = 'hello_world.txt'
  const fileType = 'text/plain'

  response.writeHead(200, {
    'Content-Disposition': `attachment; filename="${fileName}"`,
    'Content-Type': fileType,
  })

  const download = Buffer.from(fileData, 'base64')
  response.end(download)
})

app.listen(3000, () => {
  console.log('Server listening on http://localhost:3000 ...');
});
