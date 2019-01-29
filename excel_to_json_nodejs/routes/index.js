var express = require('express');
var router = express.Router();
var Excel = require('exceljs');
const multer = require('multer');
const upload = multer({ dest: 'temp/files/' });
const fs = require("fs");
const path = require('path');


/* GET home page. */
router.get('/', function (req, res, next) {
  res.render('index', { title: 'Express' });
});
// var exceltojson = require("xlsx-to-json-lc");

router.post('/import', upload.single('file'), function (req, res, next) {
  var ext;
  if(req.file!=undefined && req.file!= Error){
    var filename = req.file.originalname;
    ext = (filename.split('.')[1] ) ? filename.split('.')[1] : '';
  }
  else{
    ext = '';
  }

  //If the Incoming File is CSV 
  if (ext == 'csv') {
    let csv = require('csvtojson');
    csv()
      .fromFile(req.file.path)
      .then((result) => {
        // res.send(jsonObj);
        arr = [];
        result.map((index => {
          (index.Resource_Ldap == "" || index.Resource_Ldap == 'No Name') ? delete index : arr.push(index);
        }));

        res.send({
          StatusCode : 200,
          CountWithDummyLDAP : result.length,
          CountWithoutDummyLDAP : arr.length,
          DummyLDAPs : parseInt(result.length) - parseInt(arr.length),
          ImportData : arr
        });
      })
  }

  //If the incoming file is XLSX or XLS 
  else if (ext == 'xlsx' || ext == 'xls') {
    let exceltojson = require("xlsx-to-json");
    exceltojson({
      input: req.file.path,
      output: 'temp/files/output.json',
      // sheet: "sheetname",  // specific sheetname inside excel file (if you have multiple sheets)
      // lowerCaseHeaders:true //to convert all excel headers to lowr case in json
    }, function (err, result) {
      if (err) {
        console.log(err);
      } else {
        // console.log(result);
        arr = [];
        result.map((index => {
          (index.Resource_Ldap == "" || index.Resource_Ldap == 'No Name') ? delete index : arr.push(index);
        }))
        res.send({
          StatusCode : 200,
          CountWithDummyLDAP : result.length,
          CountWithoutDummyLDAP : arr.length,
          DummyLDAPs : parseInt(result.length) - parseInt(arr.length),
          ImportData : arr
        });
        //result will contain the overted json data
      }
    });
  }
  else{
    res.send({
      StatusCode : 400,
      Info : "Please import valid files with extension .xlsx,.xls,csv format"
    })
  }

//Delete the stored file in local project which is saved everytime while importing
  const directory = 'temp/files/';
  fs.readdir(directory, (err, files) => {
    if (err) throw err;
    for (const file of files) {
      fs.unlink(path.join(directory, file), err => {
        if (err) throw err;
      });
    }
  })

  //   // read from a file
  // var workbook = new Excel.Workbook();
  // workbook.xlsx.readFile(req.file.path)
  //     .then(function(data) {
  //         // use workbook
  //         res.send(data)
  //         // console.log(data)
  //     });
  // console.log(req.file.path)
})
module.exports = router;
