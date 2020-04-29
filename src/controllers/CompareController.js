

var parser = require('simple-excel-to-json')
const ExcelJS = require('exceljs');
const fs = require('fs')
const path = require('path')


module.exports = {

    async  compareDI(di_old, di_new, caminho, fonte) {
        var oldDI = parser.parseXls2Json(caminho + di_old)[0];
        console.log(oldDI[0])
        
        var newDI = parser.parseXls2Json(caminho+ di_new)[0];
        relatorio = []
        var relatorioExcluidos = []
        var excluido;
        var encontrado;
        
         for(var v=0; v < oldDI.length-1; v++){
            encontrado = undefined
            excluido = undefined
            encontrado = newDI.find(el => el.CableTAG == oldDI[v].CableTAG);
        
            if(encontrado == undefined){
            excluido = oldDI.find(el => el.CableTAG == oldDI[v].CableTAG)
            }
            
            if(excluido){
                
              relatorioExcluidos.push([excluido.Nitem, excluido.CableTAG, excluido.Origem, excluido.Coluna_de_Origem, excluido.Doc_Referencia_Origem,
              excluido.Regua_de_Origem, excluido.Borne_de_Origem, excluido.Condutor, excluido.Destino, excluido.Coluna_de_Destino, excluido.Doc_Referencia_Destino,
            excluido.Regua_de_Destino, excluido.Borne_de_Destino, excluido.Formacao, excluido.Codigo, excluido.Bitola, excluido.Revisao, "Excluded"])

             
             }
            
                 
       
        }
        
       
      
        for(var n=0; n<newDI.length; n++){

    await this.compara(newDI[n].CableTAG,newDI[n].Nitem, newDI[n], oldDI).then((result)=> {
        alteracao = result
        
    })
   
    if(alteracao[0] == "Equal"){
      if(newDI[n].Nitem > 1){
        for(var ab = newDI[n].Nitem ; ab > 1 ; ab = ab-1){
          if(relatorio[n-ab+1][1] == "Changed" & relatorio[n-ab+1][2] == newDI[n].Revisao){
            alteracao[2] = "Changed"
            alteracao[1] = newDI[n].Revisao
          }
        
        }
      }
      relatorio.push([newDI[n].CableTAG, alteracao[2], alteracao[1]])
    }

      if(alteracao[0] == "Changed"){
        // Ve que o cabo foi alterado na nova revisão.
       
        if(newDI[n].Nitem > 1){
          // Se o numero do item é maior que 2
          for(var ab = newDI[n].Nitem ; ab > 1 ; ab = ab-1){

            relatorio[n-ab+1][1] = "Changed"
       
            //relatorio[n -2 +1][1] = Changed
            
            relatorio[n-ab+1][2] = newDI[n].Revisao
           
            //Relatório[n -2 +1] = Nova revisão.
            //Relatorio[9]
          } 
        }
        relatorio.push([newDI[n].CableTAG, alteracao[2], alteracao[1]])
      }
     if (alteracao[0] == "Included"){
      relatorio.push([newDI[n].CableTAG, alteracao[2], alteracao[1]])

      }
     
    }
  

  
    
    this.createExcelCompara(fonte, 'comparacaoDI', relatorio, relatorioExcluidos)
    
    fs.readdir(caminho, (err, files) => {
        if (err) throw err;
      
        for (const file of files) {
          fs.unlink(path.join(caminho, file), err => {
            if (err) throw err;
          });
        }
      });

       
    },
    async  compara(TagCabo, Ncondutor, cable, cabosJson){
        for(i=0; i<cabosJson.length; i++){
       
          if(cabosJson[i].CableTAG == TagCabo && cabosJson[i].Nitem == Ncondutor){
           
           var Codigo = cabosJson[i].Codigo
           var Origem = cabosJson[i].Origem
           var Corigem = cabosJson[i].Coluna_de_Origem
           var Destino = cabosJson[i].Destino
           var Cdestino = cabosJson[i].Coluna_de_Destino
           var formacao = cabosJson[i].Formacao
           var bitola = cabosJson[i].Bitola
           var RefDocOrigem = cabosJson[i].Doc_Referencia_Origem
           var RefDocDestino = cabosJson[i].Doc_Referencia_Destino
           var ReguaOrigem  = cabosJson[i].Regua_de_Origem
           var BorneOrigem  = cabosJson[i].Borne_de_Origem
           var Condutor  = cabosJson[i].Condutor
           var ReguaDestino  = cabosJson[i].Regua_de_Destino
           var BorneDestino  = cabosJson[i].Borne_de_Destino
           var Revisao = cabosJson[i].Revisao
           var Alteracao = cabosJson[i].Alteracao
           i = cabosJson.length;
           if(
             cable.Origem != Origem || 
             cable.Codigo != Codigo  ||
             cable.Coluna_de_Origem != Corigem ||
             cable.Destino != Destino ||
             cable.Coluna_de_Destino != Cdestino ||
             cable.Formacao != formacao ||
             cable.Bitola != bitola ||
             
                 
             
             cable.Regua_de_Origem != ReguaOrigem || 
             cable.Borne_de_Origem != BorneOrigem || 
             cable.Condutor != Condutor || 
             cable.Regua_de_Destino != ReguaDestino || 
             cable.Borne_de_Destino != BorneDestino 
       ){
         
            tipoAlteracao = "Changed"
            Revisao = cable.Revisao
            Alteracao = "Changed"
            
           
           
          }else{
            tipoAlteracao = "Equal"
            Revisao  = Revisao
            Alteracao = Alteracao
            
          }
       
       
        }else{
          tipoAlteracao = "Included"
          Revisao = cable.Revisao
          Alteracao = "Included"
       
        }
       
       
       }
       if(tipoAlteracao == 'Changed'){  
    
       
       }
       return [tipoAlteracao, Revisao, Alteracao];
       
        },

        async  compareLC(LC_old, LC_new, caminho, fonte) {
            var oldLC = parser.parseXls2Json(caminho + LC_old)[0];
            
            
            var newLC = parser.parseXls2Json(caminho+ LC_new)[0];
            relatorio = []
            relatorioExcluidos = []
            

            for(var v=0; v < oldLC.length-1; v++){
              encontrado = undefined
              excluido = undefined
              encontrado = newLC.find(el => el.CableTAG == oldLC[v].CableTAG);
              
              if(encontrado == undefined){
              excluido = oldLC.find(el => el.CableTAG == oldLC[v].CableTAG)
              }
              
              if(excluido){
                  
                relatorioExcluidos.push([excluido.Codigo, excluido.CableTAG, excluido.Origem, excluido.Coluna_de_Origem, excluido.Destino,
              excluido.Coluna_de_Destino, excluido.Condutor, excluido.Formacao, excluido.Bitola, excluido.Distancia,excluido.Rota, excluido.Revisao, "Excluded"])
  
               
               }
              
                   
         
          }
            for(n=0; n<newLC.length; n++){
    
      
        await this.comparaLC(newLC[n].CableTAG, newLC[n], oldLC).then((result)=> {
            alteracao = result
           
  
            
        })
       
    
        if(alteracao[0] == "Equal"){
          relatorio.push([newLC[n].CableTAG, alteracao[2], alteracao[1]])
        }else {
          relatorio.push([newLC[n].CableTAG, alteracao[0], newLC[n].Revisao])
        }
  
         
      }
    
      
       
        this.createExcelCompara(fonte, 'comparaLC', relatorio, relatorioExcluidos)
        
        fs.readdir(caminho, (err, files) => {
            if (err) throw err;
          
            for (const file of files) {
              fs.unlink(path.join(caminho, file), err => {
                if (err) throw err;
              });
            }
          });
    
           
        },
        async  comparaLC(TagCabo, cable, cabosJson){
            for(i=0; i<cabosJson.length; i++){
           
              if(cabosJson[i].CableTAG == TagCabo){
               
               var Codigo = cabosJson[i].Codigo
               var Origem = cabosJson[i].Origem
               var Corigem = cabosJson[i].Coluna_de_Origem
               var Destino = cabosJson[i].Destino
               var Cdestino = cabosJson[i].Coluna_de_Destino
               var formacao = cabosJson[i].Formacao
               var bitola = cabosJson[i].Bitola
               var Condutor  = cabosJson[i].Condutor
               var Distancia  = cabosJson[i].Distancia
               var Rota  = cabosJson[i].Rota
               var Revisao = cabosJson[i].Revisao
               var Alteracao = cabosJson[i].Alteracao
               i = cabosJson.length;
               if(
                 cable.Origem != Origem || 
                 cable.Codigo != Codigo  ||
                 cable.Coluna_de_Origem != Corigem ||
                 cable.Destino != Destino ||
                 cable.Coluna_de_Destino != Cdestino ||
                 cable.Formacao != formacao ||
                 cable.Bitola != bitola ||
                 cable.Distancia != Distancia ||
                 cable.Rota != Rota||
                 cable.Condutor != Condutor  
           ){
             
                tipoAlteracao = "Changed"
                
               
               
              }else{
                tipoAlteracao = "Equal"
                
              }
           
           
            }else{
              tipoAlteracao = "Included"
            }
           
           
           }
           
           //return tipoAlteracao;
           return [tipoAlteracao, Revisao, Alteracao];
           
            },

        async  cabosExcluidos(TagCabo, newDI, oldDI){
          
          encontrado = newDI.find(el => el.CableTAG == TagCabo);
        
          if(encontrado == undefined){
          excluido = oldDI.find(el => el.CableTAG == TagCabo)
          return(excluido)
          
          }
            },

    async createExcelCompara(diretorio, excelName, vetor, relatorioExcluidos){
            var wb = new ExcelJS.Workbook();
            var ws = wb.addWorksheet(excelName);
            var wsExcluded = wb.addWorksheet("CabosExcluidos");
            
      
           
               
            
            var compare =  vetor;
            
            
            for(var a = 2; a< compare.length; a++){
                
                var row = ws.getRow(a)
                
                row.getCell(1).value = compare[a-2][0];
              
                row.getCell(2).value = compare[a-2][1];
                row.getCell(3).value = compare[a-2][2];

               
            }
          /*  if(relatorioExcluidos.length > 0){
              for(var b = 2; b <relatorioExcluidos.length; b++){
                var rowExcluded = wsExcluded.getRow(b)
                rowExcluded.getCell(1).value = relatorioExcluidos[b-2][0]
                rowExcluded.getCell(2).value = relatorioExcluidos[b-2][1]
                rowExcluded.getCell(3).value = relatorioExcluidos[b-2][2]
                rowExcluded.getCell(4).value = relatorioExcluidos[b-2][3]
                rowExcluded.getCell(5).value = relatorioExcluidos[b-2][4]
                rowExcluded.getCell(6).value = relatorioExcluidos[b-2][5]
                rowExcluded.getCell(7).value = relatorioExcluidos[b-2][6]
                rowExcluded.getCell(8).value = relatorioExcluidos[b-2][7]
                rowExcluded.getCell(9).value = relatorioExcluidos[b-2][8]
                rowExcluded.getCell(10).value = relatorioExcluidos[b-2][9]
                rowExcluded.getCell(11).value = relatorioExcluidos[b-2][10]
                rowExcluded.getCell(12).value = relatorioExcluidos[b-2][11]
                rowExcluded.getCell(13).value = relatorioExcluidos[b-2][12]
                rowExcluded.getCell(14).value = relatorioExcluidos[b-2][13]
                rowExcluded.getCell(15).value = relatorioExcluidos[b-2][14]
                rowExcluded.getCell(16).value = relatorioExcluidos[b-2][15]
                rowExcluded.getCell(17).value = relatorioExcluidos[b-2][16]
                rowExcluded.getCell(18).value = relatorioExcluidos[b-2][17]
                rowExcluded.getCell(19).value = relatorioExcluidos[b-2][18]



              }

            } */
            await wb.xlsx.writeFile(diretorio + '/public/' + excelName + '.xlsx')
        .then(function(){
            console.log('done') 
          

      });
          
          

}
}



  

