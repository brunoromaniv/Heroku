var express = require('express');
var router = express.Router();
const fs = require('fs');
const path = require('path');


router.get('/downloadLC', (req, res) => {
    
    var file2 = fs.createReadStream(path.join(__dirname , "../",'/public/', 'comparacaoLC.xlsx'));
    file2.pipe(res);
   
  
});

module.exports = router;