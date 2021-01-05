const fs = require("fs");
const { google } = require('googleapis');

const config = require('./config');

function upload(fileName, filePath, callback) {
  require("./gdrive-auth")((auth) => {
    const fileMetadata = {
      name: fileName,
      parents: [config.folderId]
    };

    const media = {
      mimeType: "text/csv",
      body: fs.createReadStream(filePath)
    }

    const drive = google.drive({ version: 'v3', auth });
    drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: 'id'
    }, function (err, file) {
      if (err) {
        // Handle error
        console.error(err);
      } else {
        callback(file.data.id);
      }
    });
  });
}

function download() {
  return new Promise((resolve, reject) => {

    require("./gdrive-auth")((auth) => {
      const drive = google.drive({ version: 'v3', auth }); // Authenticating drive API

      var fileId = '1RHouD2Lq3dqKReQiiNX1LMYiZYgX3WQ5ta4NeVttEM4';
      var dest = fs.createWriteStream(`${__dirname}\\base\\Alocação e produtividade das equipes.xlsx`);
      drive.files.export({
        fileId: fileId,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }, { responseType: 'stream' }, function (err, response) {
        response.data
          .on('end', function () {
            console.log("Download do arquivo base finalizado");
            resolve();
          })
          .on('error', function (err) {
            console.log('Error during download', err);
            return process.exit();
          })
          .pipe(dest);
      });


    })

  })
}

function del(fileId) {

  return new Promise((resolve, reject) => {

    require("./gdrive-auth")((auth) => {

      const drive = google.drive({ version: 'v3', auth }); // Authenticating drive API

      // Deleting the image from Drive
      drive.files
        .delete({
          fileId: fileId,
        })
        .then(
          async function (response) {

            if (response.status == 204) return "arquivo apagado";
            else console.log(response);

            resolve(response);
          },
          function (err) {

            console.log('err', err);
            reject(err);
          }
        );
    });

  })
}

/**
 * List files in past
 */
async function searchFile(name, _delete) {

  _delete = _delete || true;

  // Autentica
  require("./gdrive-auth")((auth) => {

    // Intancia a api
    const drive = google.drive({ version: 'v3', auth });

    return new Promise((resolve, reject) => {

      // Lista os arquivos dentro da pasta
      drive.files.list({
        q: `'${config.folderId}' in parents`,
        fields: 'nextPageToken, files(id, name), files/parents',
      }, (err, res) => {
        if (err) return console.log('The API returned an error: ' + err);

        // Pega os arquivos que foram retornados
        const files = res.data.files;
        let findFiles = [];

        // Checa se teve retorno
        if (files.length) {

          console.log('Files:');
          files.map(async (file) => {
            console.log(`${file.name} (${file.id})`);



            // verifica se existe algum arquivo com o mesmo nome
            if (file.name == name) {

              findFiles.push(file);

              // Verifica se é para deletar
              if (_delete) {
                console.log("deletando arquivo anterior");
                await del(file.id);
                console.log("arquivo deletado");
              }

            }
          });
        } else {
          console.log('No files found.');
        }

        // Retorna o resultada da busca para
        resolve(findFiles);

      });

    });

  })
}

module.exports = { upload: upload, del: del, searchFile: searchFile, download: download };