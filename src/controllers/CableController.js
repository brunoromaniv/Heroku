const Cable = require('../models/Cables')

module.exports = {
    async store(request, response) {
        //console.log(request.body);
        
        
        const {TAG, Codigo, Origem, Corigem, Destino, Cdestino, formacao, bitola, RefDocOrigem, RefDocDestino, Fornecedor, ReguaOrigem, BorneOrigem, Condutor, ReguaDestino, BorneDestino} = request.body;
       
        
       let cable = await Cable.findOne({TAG})
       
       
       
       if (cable){
           console.log(true)
          cable = await Cable.updateMany({
            TAG, 
            Codigo, 
            Origem, Corigem, 
            Destino, 
            Cdestino, 
            formacao, 
            bitola, 
            RefDocOrigem, 
            RefDocDestino, 
            Fornecedor, 
            ReguaOrigem, 
            BorneOrigem, 
            Condutor, 
            ReguaDestino, 
            BorneDestino,
            $inc: {Revisao: 1}
           })
           
       } else {

        cable = await Cable.create ({
           TAG, 
           Codigo, 
           Origem, Corigem, 
           Destino, 
           Cdestino, 
           formacao, 
           bitola, 
           RefDocOrigem, 
           RefDocDestino, 
           Fornecedor, 
           ReguaOrigem, 
           BorneOrigem, 
           Condutor, 
           ReguaDestino, 
           BorneDestino,
           Revisao: 0
       })
    }
       return response.json(cable)
    }
};