var parser = require('simple-excel-to-json')
const ExcelJS = require('exceljs');
const fs = require('fs')
const path = require('path')
const Graph = require('node-dijkstra')
var teste;

module.exports = {

  async transformaXLSX(arquivo, caminho, caminhoJson) {
        var viasJson = await parser.parseXls2Json(caminho + arquivo)[0];
    fs.readdir(caminho, (err, files) => {
      if (err) throw err;

   /*    for (const file of files) {
      
        fs.unlink(path.join(caminho, file), err => {
          if (err) throw err;
        }); 
      
    }*/
      fs.unlink(caminho + arquivo, function(err){
        if(err && err.code == "ENOENT"){
          return err
        } else if (err){
          return err
        }else {
          
        }
      })
    });
    this.converteValores(viasJson)
    var verificacao = this.testFile(viasJson)
    var resultado;
    await verificacao.then(function (val) {
      resultado = val
      
    })
    fs.writeFileSync(caminhoJson + 'vias.json', JSON.stringify(viasJson))
    var resultado = {
      vias: viasJson,
      teste: resultado
    }

    return resultado

  },
  async converteValores(file){
    for(var i=0; file.length > i; i++){
      file[i].COLUNA = (file[i].COLUNA).toString();
      file[i].COMPRIMENTO = parseFloat(file[i].COMPRIMENTO);
      file[i].SECAO = parseFloat(file[i].SECAO);
    }
    
  },

  async testFile(file) {
    var verificaPaineleColuna = [];

    for (var i = 0; file.length > i; i++) {
      var painel = file[i].PAINEL
      var coluna = file[i].COLUNA
      var via = file[i].VIA
      var classVia = via.substr(6, 1)
      var classificacao = file[i].CLASSIFICACAO

      var comprimento = file[i].COMPRIMENTO
      var secao = file[i].SECAO
      var conexoes = file[i].CONEXOES
      var conexoesSeparadas = conexoes.split(';')
      var colunasSeparadas = coluna.split(';')
      var paineisSeparados = painel.split(';')
      if (classificacao == undefined || comprimento == undefined || secao == undefined || conexoes == undefined) {
        verificaPaineleColuna.push("Favor verificar as colunas se estão sem caracteres especiais.")
        return verificaPaineleColuna
      }


      for (var w = 0; colunasSeparadas.length > w; w++) {
        if (colunasSeparadas[w] == "" && paineisSeparados[w] != "") {
          verificaPaineleColuna.push("Error: " + via + ": Painel cadastrado sem coluna");
        }
        if (colunasSeparadas[w] != "" && paineisSeparados[w] == "") {
          verificaPaineleColuna.push("Error: " + via + ": Coluna cadastrada sem painel");
        }
      }

      if (via.indexOf("-") != -1 && via.length != 14) {
        verificaPaineleColuna.push("Error: " + via + ": Via cadastrada está com nomenclatura errada");
      } else if (via.indexOf("-") == -1) {
        verificaPaineleColuna.push("Warning: " + via + ": Possui nome diferente do padrão, favor verificar.")
      } else if (classVia != classificacao && classVia != "N" && classificacao != "D" && via.substr(5, 2) != "CP") {

        verificaPaineleColuna.push("Error: " + via + ": Via cadastrada é diferente de sua classificação");

      }
      if (comprimento == 0) {
        verificaPaineleColuna.push("Error: " + via + ": Comprimento da Via está com valor nulo");
      }
      if (comprimento < 0.4) {
        verificaPaineleColuna.push("Warning: " + via + ": Favor verificar o comprimento cadastrado para essa via");
      }
      if (secao == 0) {
        verificaPaineleColuna.push("Error: " + via + ": Via sem seção cadastrada")
      }
      for (var j = 0; conexoesSeparadas.length > j; j++) {
        if (via == conexoesSeparadas[j]) {
          verificaPaineleColuna.push("Error: " + via + ": A via possui em suas conexões ela mesma")
        }
      }
    }

    return verificaPaineleColuna.sort()

  },

  async transformaXLSXdePara(arquivo, caminho, caminhoJson, caminhoRaiz) {

    var dePARA = await parser.parseXls2Json(caminho + arquivo)[0];
    fs.readdir(caminho, (err, files) => {
      
      for (const file of files) {
       
       
        fs.unlink(path.join(caminho, file), err => {
          if (err) throw err;
        });
       
      }
      
    });


    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet("DE_PARA");


    for (var i = 0; i < dePARA.length; i++) {
      var menorCaminho = this.shortestPath(dePARA[i].DE, dePARA[i].PARA, dePARA[i].BANDEJA)
      
      var row = ws.getRow(i + 1)
      await menorCaminho.then(function (result) {

        row.getCell(1).value = dePARA[i].DE

        row.getCell(2).value = dePARA[i].PARA
        if (result.path != null) {
          row.getCell(3).value = result.path

          row.getCell(4).value = result.cost
        } else {
          row.getCell(3).value = "Rota não encontrada"

          row.getCell(4).value = result.cost
        }
      })
    }
    await wb.xlsx.writeFile(caminhoRaiz + '/public/' + "dePARA" + '.xlsx')
    return dePARA

  },

    async shortestPath(origem, destino, classificacao){
      const route = new Graph()
      let vias =await  require('../../uploads/viasJson/vias.json')
    
      for(var i=0; i < vias.length; i++){
        if(vias[i].CLASSIFICACAO == classificacao){
        var conexoes = vias[i].CONEXOES.split(';')
        var conec = {}
        var conecV = {}
        var via = vias[i].VIA
        var paineis = vias[i].PAINEL.split(';')
        var dist = vias[i].COMPRIMENTO
        var colunas = vias[i].COLUNA

        if (colunas != undefined) {
          var colunas = ((vias[i].COLUNA).toString()).split(';')
        }
        for (j = 0; j < conexoes.length; j++) {
          var conexao = conexoes[j]

          conec[conexao] = dist
          conecV[conexao] = dist
        }
        for (v = 0; v < paineis.length; v++) {
          var painel = paineis[v]
          var coluna = colunas[v]

          if (painel != "" && (coluna != undefined || coluna != "")) {

            conecV[painel + '-' + coluna] = dist
            if (painel + '-' + coluna == origem) {
              var ViaConectada = via
            }
            route.addNode(via, conecV)
            conec[via] = dist
            if (conec != "") {
              route.addNode(painel + "-" + coluna, conec)

            }

          } else if (painel != "" && coluna == "") {
            conecV[painel] = dist
            route.addNode(via, conecV)


            conec[via] = dist
            route.addNode(painel, conec)

          }
          else {

            route.addNode(via, conec)

          }
        }
      }
    }

  
      
    
      console.log(origem, destino)
        var a =0
        var b= 0
        var c = 0
       for(i=0; i < vias.length; i++){
       a = vias[i].PAINEL.split(";");
       b = vias[i]
       c = (vias[i].COLUNA).toString().split(";");
      
        for(j=0; j < a.length; j++ ){
          if(a[j] + "-" + c[j]  == origem){
            var OrigemSemColuna = b.PAINEL.indexOf(origem.substr(0, origem.lastIndexOf("-")))
            var DestinoSemColuna = b.PAINEL.indexOf(destino.substr(0, destino.lastIndexOf("-")))
            console.log(OrigemSemColuna, DestinoSemColuna)
            if((OrigemSemColuna != -1) && (DestinoSemColuna!=-1) && (OrigemSemColuna !== DestinoSemColuna)){
              console.log(b)
              
              custo = parseFloat(b.COMPRIMENTO)
              var menorCaminho = {
                path: [origem, b.VIA, destino],
                cost: custo
              }
            }
          }

        }
      }
    
    if (menorCaminho == undefined) {
      var shortestPath = route.path(origem, destino, { cost: true })
      console.log(shortestPath)
      shortestPath.path.splice(0, 1)
      shortestPath.path.unshift(ViaConectada)
      shortestPath.path.unshift(origem)
      console.log('Esse é o menor caminho calculado' + shortestPath)
      return shortestPath
    }else{
      console.log('Esse é o menor caminho forçado' + menorCaminho)
      return menorCaminho
    }

    }
  }





  

