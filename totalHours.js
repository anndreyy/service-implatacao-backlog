const readXlsxFile = require('read-excel-file/node');
const base = "./base/Alocação e produtividade das equipes.xlsx";
const { convertArrayToCSV } = require('convert-array-to-csv');
var header = [
    'month', 'value', 'agrupador'
];

const execute = async () => {

    const rows = await getRows('Horas Totais');

    var newRows = [];
    var months = rows[0];
    var lastMonth = ""; // Guarda o ultimo mes no loop
    

    // console.log(rows);
    // return false;

    //loop nos meses
    monthc: for(const j in months){
        const month = months[j];

        if(j == 0)
            continue monthc;

        rowsc: for(const i in rows){ // loop nos valores
            const row = rows[i];

            if(i == 0){
                continue rowsc;
            }

            newRows.push( [
                month, // mês
                row[j], // valor
                row[0] // agrupamento
            ]);
        }
    }

    // console.log(newRows);

    saveToCSV(newRows);

}

const getSheets = () => {
    return readXlsxFile(base, { getSheets: true });
}

const getRows = (sheetName) => {
    return readXlsxFile(base, { sheet: sheetName });
}

const saveToCSV = (arr) => {
    const csvFromArrayOfArrays = convertArrayToCSV(arr, {
        header,
        separator: ';'
    });

    fs = require('fs');
    fs.writeFile('./export/totalHours.csv', csvFromArrayOfArrays,{encoding: 'ascii'}, function (err) {
        if (err) return console.log(err);

        console.log("foi: totalHours");

    });
}


// Executa
exports.execute = execute;