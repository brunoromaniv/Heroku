var express = require('express');
var router = express.Router();
const fs = require('fs');
const path = require('path');


router.get('/download', (req, res) => {
    
    var file = fs.createReadStream(path.join(__dirname , "../",'/public', 'interligacao.xlsx'));
    file.pipe(res);
   
  
});

module.exports = router;