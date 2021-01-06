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
                removeAcento(row[0]) // agrupamento
            ]);
        }
    }

    // console.log(newRows);

    await saveToCSV(newRows);

}

const removeAcento = (text) =>
{    
    if(typeof text !== 'string'){
        return text;
    }
    // text = text.toLowerCase();                                                         
    text = text.replace(new RegExp('[ÁÀÂÃ]','gi'), 'a');
    text = text.replace(new RegExp('[ÉÈÊ]','gi'), 'e');
    text = text.replace(new RegExp('[ÍÌÎ]','gi'), 'i');
    text = text.replace(new RegExp('[ÓÒÔÕ]','gi'), 'o');
    text = text.replace(new RegExp('[ÚÙÛ]','gi'), 'u');
    text = text.replace(new RegExp('[Ç]','gi'), 'c');
    return text;                 
}

const getSheets = () => {
    return readXlsxFile(base, { getSheets: true });
}

const getRows = (sheetName) => {
    return readXlsxFile(base, { sheet: sheetName });
}

const saveToCSV = async (arr) => {
    const csvFromArrayOfArrays = convertArrayToCSV(arr, {
        header,
        separator: ';'
    });

    fs = require('fs');
    fs.writeFileSync('./export/totalHours.csv', csvFromArrayOfArrays,{encoding: 'ascii'});

    console.log("foi totalHours")
}


// Executa
exports.execute = execute;