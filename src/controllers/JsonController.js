const fs = require('fs');
var parser = require('fast-xml-parser');
var he = require('he');
const ExcelJS = require('exceljs');

const numeroConexoes ={
  KDP(codigo) {
    return (parseInt(codigo.substr(4,2)) + 1)
  },
  KAD(codigo){
    return (parseInt(codigo.substr(5,1)))
  },
  KFB(codigo){
    return (parseInt(codigo.substr(6,2)))
  },
  KNB(codigo){
    return ((parseInt(codigo.substr(4,2))*3)+1)
  },
  KVI(codigo){
    return ((parseInt(codigo.substr(5,1))*2) +1)
  },
  KVF(codigo){
      return 1
  },
  KBA(codigo){
      return 1
  }

}


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
       
       return await this.createExcelJS(dirname, nome, jsonObj);
      },
       


    async createExcelJS(diretorio, excelName, json){
      var wb = new ExcelJS.Workbook();
      var ws = wb.addWorksheet('Interligacao');
      var indice = ws.getRow(1);
      var relatorio = []
      var semBorne = false;
      var semDOC = false;
      var borneOrigem = ''
      var borneDestino = ''
      var problemasDOC = 1;
      
      var listacabos =  json;
      var LISTA = [];
    
    var a = 2;
    cableNumber = listacabos.EplanLabelling.Document.Page.length;
   
    
    for (var j = 0; j < cableNumber-1; j++){
        
      var problemasDOC = 1;
        var Cabo = [];
       
        var body = listacabos.EplanLabelling.Document.Page[j].Line.length
        var numeroCedulas = 18;
        if (body == undefined){
            teste = 1
            var docREF1 = listacabos.EplanLabelling.Document.Page[j].Header.Property[8].PropertyValue
             var docREF2 = listacabos.EplanLabelling.Document.Page[j].Header.Property[9].PropertyValue
        }else
        { teste = body-1};
        var header = listacabos.EplanLabelling.Document.Page[j].Header.Property.length
        var cableTAG = listacabos.EplanLabelling.Document.Page[j].Header.Property[1].PropertyValue
        var cableCode = listacabos.EplanLabelling.Document.Page[j].Header.Property[0].PropertyValue
        var docREF1 = listacabos.EplanLabelling.Document.Page[j].Header.Property[8].PropertyValue
        var docREF2 = listacabos.EplanLabelling.Document.Page[j].Header.Property[9].PropertyValue
     
        var problemas = false;
       
        var code = '';
       

      
    for (var i = 0; i < header; i++) {
        
           
 
       var row = ws.getRow(a)
       row.getCell(i+1).value = listacabos.EplanLabelling.Document.Page[j].Header.Property[i].PropertyValue
        if(cableCode){
            if(docREF1 == '' || docREF2 == ''){
                semDOC = true;
            }
        }
       
    };
    
    let Origem = [];
    
    if (body == undefined  ){
        for(v = 0; v < 5; v++){ 

            row.getCell(v+header+1).value = listacabos.EplanLabelling.Document.Page[j].Line.Label.Property[v].PropertyValue
              
      
            };
             borneOrigem = listacabos.EplanLabelling.Document.Page[j].Line.Label.Property[0].PropertyValue
             borneDestino = listacabos.EplanLabelling.Document.Page[j].Line.Label.Property[3].PropertyValue
           if(cableCode){
               if(borneOrigem == '' || borneDestino == ''){
                    semBorne = true;
               }
           }


           
          
            
    
            for(var num=1 ; num < numeroCedulas+1; num++){
              row.getCell(num).border = {
                bottom: {style:'medium'},
                 right: {style:'thin'}
              }
            }
    } else {
     
        var t = 0;
    for(var n = 0; n < listacabos.EplanLabelling.Document.Page[j].Line.length-1; n++){
        for(v = 0; v < 5; v++){ 
            
        row = ws.getRow(a+n);
       
        row.getCell(v+header+1).value = listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[v].PropertyValue
        borneOrigem = listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[1].PropertyValue
        borneDestino = listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[4].PropertyValue
        reguaOrigem = listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[0].PropertyValue
        reguaDestino = listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[3].PropertyValue
        if  (cableCode && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH1"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH2"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH3"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH4"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH5"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH6"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH7"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH8"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH9"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH10"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH11"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "SH12"
          && listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != ""
          
          && typeof listacabos.EplanLabelling.Document.Page[j].Line[n].Label.Property[2].PropertyValue != "string"
        
          ) {
            
               if((reguaOrigem == ''&& borneOrigem != '') || (reguaDestino == '' && borneDestino != '')){
                    semBorne = true;
                    
               }
           }
       
        for (var i = 0; i < header; i++) {
        row.getCell(i+1).value = listacabos.EplanLabelling.Document.Page[j].Header.Property[i].PropertyValue
        };
        };
        
        rowFormata = ws.getRow(v+header+1)
        t = t + v;
         for(var num=1 ; num < numeroCedulas+1; num++){
          row.getCell(num).border = {
            bottom: {style:'thin'},
            right: {style:'thin'}
          }
        }
      
 
        
    };
    
    
 
   if(cableCode){
    var prefixo
    prefixo = cableCode.substr(0,3)
  
    code = numeroConexoes[prefixo](cableCode)
    if(code != (listacabos.EplanLabelling.Document.Page[j].Line.length -1)){
        
        var ab = {
          TagCabo: cableTAG,
          Problemas: ["Checar Conexões"]
  
        }
        relatorio.push(ab);
        problemas = true;
      }
    
      if((docREF2 == '' || docREF1=='' )&& problemasDOC == false){
       if (problemas == false){
        var ab = {
          TagCabo: cableTAG,
          Problemas: ["Checar documento de referência"]
  
        }
        relatorio.push(ab);
        problemas = true;
        problemasDOC = 1;
       }else{
        
        ab.Problemas.push('Checar documento de referência')
        problemas = true;
        problemasDOC = 1;
        
       }
      }
      
   }
       

        for(var num=1 ; num < numeroCedulas+1; num++){
          row.getCell(num).border = {
            bottom: {style:'medium'},
             right: {style:'thin'}
          }
        }
    };   
    if(semDOC == true ){
       if (problemas == false){
        var ab = {
          TagCabo: cableTAG,
          Problemas: ["Checar documento de referência"]
             
        }
        problemas = true;
        relatorio.push(ab);
       }else{
        
        ab.Problemas.push('Checar documento de referência')
        problemas = true;
        
       }
      }
      if(semBorne == true){
       if (problemas == false){
        var ab = {
          TagCabo: cableTAG,
          Problemas: ["Checar Reguas de Bornes"]
             
        }
        problemas = true;
        relatorio.push(ab);
       }else{
        
        ab.Problemas.push('Checar Reguas de Bornes')
        
        problemas = true;
       }
      
   }
       problemas = false;
       semDOC = false;
       semBorne = false;
        
       
     a = a + teste;

    };

  


    var CODIGO = ws.getColumn(1).values
    var TAG = ws.getColumn(2).values
    var ORIGEM = ws.getColumn(3).values
    var COLORIGEM = ws.getColumn(4).values
    var DESTINO = ws.getColumn(5).values
    var COLUNADESTINO = ws.getColumn(6).values
    var FORMACAO = ws.getColumn(7).values
    var BITOLA = ws.getColumn(8).values
    var DOCREFORIGEM = ws.getColumn(9).values
    var DOCREFDESTINO = ws.getColumn(10).values
    var FORNECEDOR = ws.getColumn(11).values
    var NOTA = ws.getColumn(12).values
    var REGUAORIGEM = ws.getColumn(13).values
    var BORNEORIGEM = ws.getColumn(14).values
    var CONDUTOR = ws.getColumn(15).values
    var REGUADESTINO = ws.getColumn(16).values
    var BORNEDESTINO = ws.getColumn(17).values


    ws.getColumn(1).values = TAG
    ws.getColumn(2).values = ORIGEM
    ws.getColumn(3).values = COLORIGEM
    ws.getColumn(4).values = DOCREFORIGEM
    ws.getColumn(5).values = REGUAORIGEM
    ws.getColumn(6).values = BORNEORIGEM
    ws.getColumn(6).alignment = {vertical: 'middle', horizontal: 'center'}
    ws.getColumn(6).width = 20
    ws.getColumn(7).values = CONDUTOR
    ws.getColumn(7).alignment = {vertical: 'middle', horizontal: 'center'}
    ws.getColumn(7).width = 20
    ws.getColumn(8).values = DESTINO
    ws.getColumn(9).values = COLUNADESTINO
    ws.getColumn(10).values = DOCREFDESTINO
    ws.getColumn(11).values = REGUADESTINO
    ws.getColumn(12).values = BORNEDESTINO
    ws.getColumn(12).alignment = {vertical: 'middle', horizontal: 'center'}
    ws.getColumn(12).width = 20
    ws.getColumn(13).values = FORMACAO
    ws.getColumn(13).alignment = {vertical: 'middle', horizontal: 'center'}
    ws.getColumn(13).width = 20
    ws.getColumn(14).values = CODIGO
    ws.getColumn(15).values = BITOLA
    ws.getColumn(16).eachCell(function(cell, rowNumber){
      cell.value = '';
    })
    ws.getColumn(17).eachCell(function(cell, rowNumber){
      cell.value = '';
    })
    

      indice.getCell(1).value = "TAG";
      indice.getCell(2).value = "Origem";
      indice.getCell(3).value = "Coluna de Origem";
      indice.getCell(4).value = "Doc. Referência Origem";
      indice.getCell(5).value = "Regua de Origem";
      indice.getCell(6).value = "Borne de Origem";
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
        
      });
 

  
    
    return relatorio;


    }
}

    
    
   
    
