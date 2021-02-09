var parser = require('simple-excel-to-json')
const ExcelJS = require('exceljs');
const fs = require('fs')
const path = require('path')


module.exports = {

    async  transformaXLSX(arquivo ,caminho) {
        
        var vias = await parser.parseXls2Json(caminho + arquivo)[0];
        fs.readdir(caminho, (err, files) => {
            if (err) throw err;
          
            for (const file of files) {
              fs.unlink(path.join(caminho, file), err => {
                if (err) throw err;
              });
            }
          });
        return vias

    }
    }