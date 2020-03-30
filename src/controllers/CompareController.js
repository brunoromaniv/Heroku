

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
        relatorioExcluidos = []
        var caboExcluido = '';
        var excluido;
        console.log('chegou aqui')
        
         for(var v=0; v< oldDI.length; v++){
         // if(v>0){
            //if(oldDI[v].CableTAG != oldDI[v-1].CableTAG ){
            this.cabosExcluidos(oldDI[v].CableTAG, newDI).then((result)=> {
              caboExcluido = result
             
            });
        
        //  }else{
           // v++
          //}
           // }else{
             //this.cabosExcluidos(oldDI[v].CableTAG, newDI).then((result2)=> {
              //  caboExcluido = result2
            //  });
              if(caboExcluido == "Excluded"){
             relatorioExcluidos.push([oldDI[v].CableTAG, caboExcluido])
        console.log(caboExcluido)
            }
            
          
            
       
        }
       
      
        for(var n=0; n<newDI.length; n++){

    await this.compara(newDI[n].CableTAG,newDI[n].Nitem, newDI[n], oldDI).then((result)=> {
        alteracao = result
        
    })
   
    if(alteracao[0] == "Equal"){
      if(newDI[n].Nitem > 1){
        for(var ab = newDI[n].Nitem ; ab > 1 ; ab = ab-1){
          if(relatorio[n-ab+1][1] == "Changed"){
            alteracao[2] = "Changed"
            alteracao[1] = newDI[n].Revisao
          }
        
        }
      }
      relatorio.push([newDI[n].CableTAG, alteracao[2], alteracao[1]])
    }else {
      if(alteracao[0] == "Changed"){
        if(newDI[n].Nitem > 1){
          for(var ab = newDI[n].Nitem ; ab > 1 ; ab = ab-1){
            relatorio[n-ab+1][1] = "Changed"
            relatorio[n-ab+1][2] = newDI[n].Revisao
          }
        }
      }
     
    
     relatorio.push([newDI[n].CableTAG, alteracao[0], newDI[n].Revisao])
    }
  }

  
    console.log(relatorioExcluidos)
    this.createExcelCompara(fonte, 'comparacaoDI', relatorio)
    
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
            
           
           
          }else{
            tipoAlteracao = "Equal"
            
          }
       
       
        }else{
          tipoAlteracao = "Included"
        }
       
       
       }
       if(tipoAlteracao == 'Changed'){  
       console.log(tipoAlteracao, cable.CableTAG)
       }
       return [tipoAlteracao, Revisao, Alteracao];
       
        },

        async  compareLC(LC_old, LC_new, caminho, fonte) {
            var oldLC = parser.parseXls2Json(caminho + LC_old)[0];
            
            
            var newLC = parser.parseXls2Json(caminho+ LC_new)[0];
            relatorio = []
            relatorioExcluidos = []
            console.log('chegou aqui')
            for(n=0; n<newLC.length; n++){
    
      
        await this.comparaLC(newLC[n].CableTAG, newLC[n], oldLC).then((result)=> {
            alteracao = result
            console.log(alteracao)
  
            
        })
       
    
        if(alteracao[0] == "Equal"){
          relatorio.push([newLC[n].CableTAG, alteracao[2], alteracao[1]])
        }else {
          relatorio.push([newLC[n].CableTAG, alteracao[0], newLC[n].Revisao])
        }
        console.log(alteracao)
         
      }
    
      
       
        this.createExcelCompara(fonte, 'comparaLC', relatorio)
        
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

        async  cabosExcluidos(TagCabo, cabosJson){
            for(i=0; i<cabosJson.length-1; i++){
           
              if(cabosJson[i].CableTAG == TagCabo){
               
                i = cabosJson.length
                tipoAlteracao = "Equal"
                
               
              }else{
                tipoAlteracao = "Excluded"
              }
           
           
            }
                  
            
          //console.log(TagCabo, tipoAlteracao)
           return tipoAlteracao
         
            },

    async createExcelCompara(diretorio, excelName, vetor){
            var wb = new ExcelJS.Workbook();
            var ws = wb.addWorksheet(excelName);
      
           
               
            
            var compare =  vetor;
            console.log(compare[1])
            
            for(var a = 2; a< compare.length; a++){
                
                var row = ws.getRow(a)
                
                row.getCell(1).value = compare[a-2][0];
              
                row.getCell(2).value = compare[a-2][1];
                row.getCell(3).value = compare[a-2][2];

               
            }
            await wb.xlsx.writeFile(diretorio + '/public/' + excelName + '.xlsx')
        .then(function(){
            console.log('done')
          

      });
          
          

}
}



  

