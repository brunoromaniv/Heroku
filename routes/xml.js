var express = require('express');
var router = express.Router();
const JsonController = require('../src/controllers/JsonController2');
const path = require('path');
const multer = require('multer');

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, "./uploads");

    },
    filename: (req, file, cb)=>{
        const {name, ext} = path.parse(file.originalname);
        cb(null, "Xml")
        //cb(null, `${name}-${Date.now()}.${ext}`)
    }
});

const upload = multer({storage});

router.post('/xml', upload.single("file"), async (req, res)=>{
    console.log(req.body, req.files);
    console.log(path.join(__dirname , "../"))
    //console.log(path.join(__dirname , "../uploads"))
     await   JsonController.createJson(path.join(__dirname , "../uploads", 'Xml'), "Xml2", path.join(__dirname , "../"))
      
        
   res.send('ok')
   
   
});

module.exports = router;