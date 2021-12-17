const express = require('express');
const mongoose = require('mongoose');
const path = require('path');
const app = express();
const cors = require('cors');


const baseDir = path.join(__dirname, 'build');
app.use(express.json())

/* mongoose.connect('mongodb+srv://brunoromaniv:brunoromaniv@cluster0-ixmuf.mongodb.net/EplantoExcel?retryWrites=true&w=majority', {
    useNewUrlParser: true,
    useUnifiedTopology: true
}); */
app.get('/*', (req,res) => {
    res.sendfile(path.join(__dirname, 'build', 'index.html'))
})
app.use(cors())

app.use(express.static(`${baseDir}`));



app.listen(process.env.PORT || 3333);
