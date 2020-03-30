const fs = require('fs');
var parser = require('fast-xml-parser');
var he = require('he');
const ExcelJS = require('exceljs');


    module.exports = {



      async createJson(origem, nome, dirname) {
        let xmlData = fs.readFileSync(origem, "utf8");
        var options = {
          attributeNamePrefix : "@_",
          attrNodeName: "attr", //default is 'false'
          textNodeName : "#text",
          ignoreAttributes : true,
          ignoreNameSpace : false,
          allowBooleanAttributes : false,
          parseNodeValue : true,
          parseAttributeValue : false,
          trimValues: true,
          cdataTagName: "__cdata", //default is 'false'
          cdataPositionChar: "\\c",
          parseTrueNumberOnly: false,
          arrayMode: false, //"strict"
          attrValueProcessor: (val, attrName) => he.decode(val, {isAttributeValue: true}),//default is a=>a
          tagValueProcessor : (val, tagName) => he.decode(val), //default is a=>a
          stopNodes: ["parse-me-as-string"]
      };
      if( parser.validate(xmlData) === true) { //optional (it'll return an object in case it's not valid)
    var jsonObj = parser.parse(xmlData,options);
 
}        
       
        await this.createExcelJS(dirname, nome, jsonObj);
      },
       


    async createExcelJS(diretorio, excelName, json){
      var wb = new ExcelJS.Workbook();
      var ws = wb.addWorksheet('Interligacao');
      var indice = ws.getRow(1);
      

      
      var listacabos =  json;
      var LISTA = [];
    
    var a = 2;
    cableNumber = listacabos.EplanLabelling.Document.Page.length;
   
    
    for (var j = 0; j < cableNumber; j++){
        
        
        var Cabo = [];
       
        var body = listacabos.EplanLabelling.Document.Page[j].Line.length
        var numeroCedulas = 18;
        if (body == undefined){
            teste = 1
        }else
        { teste = body};
        var header = listacabos.EplanLabelling.Document.Page[j].Header.Property.length
    
       

      
    for (var i = 0; i < header; i++) {
        
           
 
       var row = ws.getRow(a)
       row.getCell(i+1).value = listacabos.EplanLabelling.Document.Page[j].Header.Property[i].PropertyValue
     
       
    };
    
    let Origem = [];
    
    if (body == undefined  ){
        for(v = 0; v < 5; v++){ 

            row.getCell(v+header+1).value = listacabos.EplanLabelling.Document.Page[j].Line.Label.Property[v].PropertyValue
           
            };
           
            
    
            for(var num=1 ; num < numeroCedulas+1; num++){
              row.getCell(num).border = {
                bottom: {style:'medium'},
              }
            }
    } else {
     
        var t = 0;
    for(var n = 0; n < listacabos.EplanLabelling.Document.Page[j].Line.length; n++){
          for(v = 0; v < 5; v++){ 
            
        row = ws.getRow(a+n);
       
        row.getCell(v+header+1).value = listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[v].PropertyValue
       
        for (var i = 0; i < header; i++) {
        row.getCell(i+1).value = listacabos.EplanLabelling.Document.Page[j].Header.Property[i].PropertyValue
        };
        };

        
        
        rowFormata = ws.getRow(v+header+1)
        t = t + v;
      
 
        
    };
    
       

        for(var num=1 ; num < numeroCedulas+1; num++){
          row.getCell(num).border = {
            bottom: {style:'medium'},
          }
        }
    };   
        
       
     a = a + teste;

    };

  


    var a1 = ws.getColumn(1).values
    var a2 = ws.getColumn(2).values
    var a3 = ws.getColumn(3).values
    var a4 = ws.getColumn(4).values
    var a5 = ws.getColumn(5).values
    var a6 = ws.getColumn(6).values
    var a7 = ws.getColumn(7).values
    var a8 = ws.getColumn(8).values
    var a9 = ws.getColumn(9).values
    var a10 = ws.getColumn(10).values
    var a11 = ws.getColumn(11).values
    var a12 = ws.getColumn(12).values
    var a13 = ws.getColumn(13).values
    var a14 = ws.getColumn(14).values
    var a15 = ws.getColumn(15).values
    var a16 = ws.getColumn(16).values
    var a17 = ws.getColumn(17).values

    ws.getColumn(1).values = a2
    ws.getColumn(2).values = a3
    ws.getColumn(3).values = a4
    ws.getColumn(4).values = a9
    ws.getColumn(5).values = a12
    ws.getColumn(6).values = a13
    ws.getColumn(6).alignment = {vertical: 'middle', horizontal: 'center'}
    ws.getColumn(6).width = 20
    ws.getColumn(7).values = a14
    ws.getColumn(7).alignment = {vertical: 'middle', horizontal: 'center'}
    ws.getColumn(7).width = 20
    ws.getColumn(8).values = a5
    ws.getColumn(9).values = a6
    ws.getColumn(10).values = a10
    ws.getColumn(11).values = a15
    ws.getColumn(12).values = a16
    ws.getColumn(12).alignment = {vertical: 'middle', horizontal: 'center'}
    ws.getColumn(12).width = 20
    ws.getColumn(13).values = a7
    ws.getColumn(13).alignment = {vertical: 'middle', horizontal: 'center'}
    ws.getColumn(13).width = 20
    ws.getColumn(14).values = a1
    ws.getColumn(15).values = a8
    ws.getColumn(16).eachCell(function(cell, rowNumber){
      cell.value = '';
    })
    

      indice.getCell(1).value = "TAG";
      indice.getCell(2).value = "Origem";
      indice.getCell(3).value = "Coluna de Origem";
      indice.getCell(4).value = "Doc. Referência Origem";
      indice.getCell(5).value = "Regua de Origem";
      indice.getCell(6).value = "Regua de Origem";
      indice.getCell(7).value = "Condutor";
      indice.getCell(8).value = "Destino";
      indice.getCell(9).value = "Coluna de Destino";
      indice.getCell(10).value = "Doc. Referência Destino";
      indice.getCell(11).value = "Regua de Destino";
      indice.getCell(12).value = "Borne de Destino";
      indice.getCell(13).value = "Formação";
      indice.getCell(14).value = "Código";
      indice.getCell(15).value = "Bitola";
      indice.getCell(16).value = "Nota";

    
    await wb.xlsx.writeFile(diretorio + '/public/interligacao.xlsx')
      .then(function(){
        console.log('done')
      });
 

  
    console.log('aqui também');


    }
}

    
    
   
    
