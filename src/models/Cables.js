const mongoose = require('mongoose');

const CableSchema = new mongoose.Schema({
     
    TAG: String,
    Codigo: String,
    Origem: String,
    Corigem: String,
    Destino: String,
    Cdestino: String,
    formacao: String,
    bitola: String,
    RefDocOrigem: String,
    RefDocDestino: String,
    Fornecedor: String,
    ReguaOrigem: String,
    BorneOrigem: String,
    Condutor: String,
    ReguaDestino: String,
    BorneDestino: String,
    Revisao: Number
        

});

module.exports = mongoose.model('Cable', CableSchema);