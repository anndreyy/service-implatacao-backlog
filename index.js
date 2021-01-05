const months = require('./months');
const totalHours = require('./totalHours');

const gdrive = require("./gdrive");
const config = require('./config');


const entry = async () => {
    // baixa a base
    await gdrive.download();

    // Gera os valores por mês
    await months.execute();
    await totalHours.execute();

    await upload('months');
    await upload('totalHours');
}

// upload files to google drive
const upload = async (filename) => {
    return new Promise( async (resolve, reject) => {
        filename = `${filename}.csv`;
        var pathWithFilename = `${config.path}\\${filename}`;
    
        // Primeiro procura arquivos com o mesmo nome para apagar
        await gdrive.searchFile(filename);
    
        gdrive.upload(filename, pathWithFilename, (id) => {
            console.log("Finalizado id: ", id);

            resolve();
        })
    })   
}

// entry();

const express = require('express')
const app = express()
const port = config.port;

app.get('/', async (req, res) => {
    await entry();
    res.send('Processamento finalizado')
})

app.listen(port, () => {
  console.log(`Service implantacao backlog running http://localhost:${port}`)
})