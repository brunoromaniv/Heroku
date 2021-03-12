var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');
const mongoose = require('mongoose');
const cors = require('cors');
const fs = require('fs');
var bodyParser = require('body-parser')




const CompareController = require('./src/controllers/CompareController')
const JsonController = require('./src/controllers/JsonController');
const multer = require('multer');
var destino = path.join(__dirname, './uploads', 'Xml2.json');
const GrafoController = require('./src/controllers/GrafoController')



var indexRouter = require('./routes/index');
var usersRouter = require('./routes/users');
var xmlRouter = require('./routes/xml');
var downloadRouter = require('./routes/download');
var downloadDIRouter = require('./routes/downloadDI');
var downloadLCRouter = require('./routes/downloadLC')

var app = express();

mongoose.connect('mongodb+srv://brunoromaniv:brunoromaniv@cluster0-ixmuf.mongodb.net/EplantoExcel?retryWrites=true&w=majority', {
    useNewUrlParser: true,
    useUnifiedTopology: true
});
// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'jade');
//teste de versão

app.use(logger('dev'));
//app.use(express.json());
app.use(bodyParser.json({limit: '50mb'}))
app.use(bodyParser.urlencoded({limit: '50mb', extended: true}))
//app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'build')));
app.use(cors('http://localhost:3333'))

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
      cb(null, "./uploads");

  },
  filename: (req, file, cb)=>{
      const {name, ext} = path.parse(file.originalname);
      cb(null, "Xml")
      //cb(null, `${name}-${Date.now()}.${ext}`)
  },
  limits: {
    fieldSize: 50 * 1024 * 1024,
   
  },

 

});


const storageM = multer.diskStorage({
	destination:function(req,file,cb){
		cb(null,"./uploads/compareDI")
	},
	filename: function(req,file,cb){
		cb(null,file.originalname);
 
	}
});
const storageMX = multer.diskStorage({
	destination:function(req,file,cb){
		cb(null,"./uploads/xlsx")
	},
	filename: function(req,file,cb){
		cb(null,file.originalname);
 
	}
});


var oldDI = ''
var newDI = ''
const uploadM = multer({storage: storageM})
app.post('/comparaDI', uploadM.any(), function(req,res){
  var file = req.files;
  
  oldDI  = file[0].originalname
  newDI =  file[1].originalname
  pathA = __dirname + '/uploads/compareDI/'
  pathO = __dirname;
  CompareController.compareDI(oldDI, newDI, pathA, pathO)
 

  res.end();
});


const uploadMX = multer({storage: storageMX})
app.post('/grafoVias', uploadMX.any(), async function(req,res){
  var tamanho = req.files.length
  var vias = req.files[tamanho-1].originalname;
  

  var caminho = __dirname + '/uploads/xlsx/'
  var caminhoJson = __dirname + "/uploads/viasJson/"

  testes = await GrafoController.transformaXLSX(vias, caminho, caminhoJson )
 
  res.json(testes)
})

app.post('/shortestPath', async function(req,res){
   
    var origem = req.body.origem
    var destino = req.body.destino
    var classificacao = req.body.classificacao
  console.log(classificacao)


    var shortestPath = await GrafoController.shortestPath(origem, destino, classificacao)
    console.log(shortestPath)
  res.json(shortestPath)
})

app.post('/grafoViasDePara',uploadMX.any(), async function(req,res){

  var tamanho = req.files.length
  var vias = req.files[tamanho-1].originalname;
  

  var caminho = __dirname + '/uploads/xlsx/'
  var caminhoRaiz = __dirname
  var caminhoJson = __dirname + "/uploads/viasJson/"

  testes = await GrafoController.transformaXLSXdePara(vias, caminho, caminhoJson, caminhoRaiz )

  res.json("resposta: ok")
  //console.log(testes)
//  res.json(testes)
})

app.get('/grafoViasDeParaDownload', (req, res) => {
    
  var file = fs.createReadStream(path.join(__dirname , '/public', 'dePARA.xlsx'));
  
  file.pipe(res);
 

});


const storageML = multer.diskStorage({
	destination:function(req,file,cb){
		cb(null,"./uploads/compareLC")
	},
	filename: function(req,file,cb){
		cb(null,file.originalname);
 
	}
});

var oldLC = ''
var newLC = ''
const uploadML = multer({storage: storageML})
app.post('/comparaLC', uploadML.any(), function(req,res){
  var file = req.files;
  
  oldLC  = file[0].originalname
  newLC =  file[1].originalname
  pathA = __dirname + '/uploads/compareLC/'
  pathO = __dirname;
  CompareController.compareLC(oldLC, newLC, pathA, pathO)
 

  res.end();
})

const upload = multer({storage});

app.post('/xml', upload.single("file"), async (req, res)=>{

  //console.log(path.join(__dirname , "../uploads"))
  var relatorio =  await   JsonController.createJson(path.join(__dirname , "./uploads", 'Xml'), "Xml2", path.join(__dirname))
    
 
      
 res.json(relatorio)
 relatorio = ''
 
 
});



app.get('/download', (req, res) => {
    
  var file = fs.createReadStream(path.join(__dirname , '/public', 'interligacao.xlsx'));
  file.pipe(res);
 

});
app.get('/downloadDI', (req, res) => {
    
  var file = fs.createReadStream(path.join(__dirname , '/public', 'comparacaoDI.xlsx'));
  file.pipe(res);
 

});
app.get('/downloadLC', (req, res) => {
    
  var file2 = fs.createReadStream(path.join(__dirname , '/public', 'comparaLC.xlsx'));
  file2.pipe(res);
 

});


app.get('/ModeloLC', (req, res) => {
    
  var file2 = fs.createReadStream(path.join(__dirname , '/public', 'ModeloLC.xlsx'));
  file2.pipe(res);
 

});

app.get('/ModeloDI', (req, res) => {
    
  var file2 = fs.createReadStream(path.join(__dirname , '/public', 'ModeloDI.xlsx'));
  file2.pipe(res);
 

});






app.use('/', indexRouter);
app.use('/users', usersRouter);
//app.use('/xml', xmlRouter);
app.use('/download', downloadRouter);
app.use('/downloadDI', downloadDIRouter);
app.use('/downloadLC', downloadLCRouter);


// catch 404 and forward to error handler


// error handler






module.exports = app;



