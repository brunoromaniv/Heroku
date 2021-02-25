var parser = require('simple-excel-to-json')
const ExcelJS = require('exceljs');
const fs = require('fs')
const path = require('path')
const Graph = require('node-dijkstra')
var teste;

module.exports = {
  

    async  transformaXLSX(arquivo ,caminho, caminhoJson) {
        
        var vias = await parser.parseXls2Json(caminho + arquivo)[0];
        fs.readdir(caminho, (err, files) => {
            if (err) throw err;
          
            for (const file of files) {
              fs.unlink(path.join(caminho, file), err => {
                if (err) throw err;
              });
            }
          });
          
          fs.writeFileSync(caminhoJson + 'vias.json', JSON.stringify(vias))
      
          
         
        return vias

    },

    async shortestPath(origem, destino, classificacao){
      const route = new Graph()
      let vias = require('../../uploads/viasJson/vias.json')
     console.log(classificacao)
      for(var i=0; i < vias.length; i++){
        if(vias[i].CLASSIFICACAO == classificacao){
        var conexoes = vias[i].CONEXOES.split(';')
        var conec = {}
        var conecV = {}
        var via = vias[i].VIA
        var paineis = vias[i].PAINEL.split(';')
        var dist = vias[i].COMPRIMENTO
        var colunas = vias[i].COLUNA
        
        if(colunas != undefined){
        var colunas = vias[i].COLUNA.split(';')
      }
        
       
        for(j=0; j < conexoes.length; j++){
          var conexao = conexoes[j]
          
          conec[conexao] = dist
          conecV[conexao] = dist
          
        }
        for(v=0; v < paineis.length; v++){
        var painel = paineis[v]
        var coluna = colunas[v]
       
        if(painel != "" && (coluna != undefined || coluna != "")){
          
          conecV[painel+'-'+coluna] = dist
          route.addNode(via, conecV)
          
       
        conec[via] = dist
        if(conec != ""){
        route.addNode(painel+"-"+coluna, conec)
      
      }
        
      }else if(painel != "" && coluna ==""){
        conecV[painel] = dist
        route.addNode(via, conecV)
       
     
      conec[via] = dist
      route.addNode(painel, conec)
     
     
      }    
      else{
        
        route.addNode(via, conec)
       
       
      }
    }
        

  }
      }
      console.log(origem, destino)
      var shortestPath = route.path(origem, destino, {cost: true})
      return shortestPath
      


    }
    }