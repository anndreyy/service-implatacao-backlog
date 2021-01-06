const readXlsxFile = require('read-excel-file/node');
const { convertArrayToCSV } = require('convert-array-to-csv');

const base = "./base/Alocação e produtividade das equipes.xlsx";

var header = [
    'Equipe de Implantacao',
    'Backlog inicial (saldo em horas)',
    'Horas executadas no mes',
    'Horas adicionadas no mes',
    'Horas de vendas no mes',
    'Backlog final (saldo em horas)',
    'Agrupador',
    'data'
];



const execute = async () => {
    
    // File path.
    const arrSheets = await getSheets();
    const arrSheetMonths = getMonths(arrSheets);
    const arrData = [];

    for (const i in arrSheetMonths) {

        arrSheetMonths[i].data = await getRows({ name: arrSheetMonths[i].name })
        arrData.push(formatedData(arrSheetMonths[i].data, arrSheetMonths[i].name));
    }

    saveToCSV(arrData.flat());
};



const getSheets = () => {
    return readXlsxFile(base, { getSheets: true });
}

const getMonths = (sheets) => {

    return sheets.filter(sheet => {
        return sheet.name.search(/[.]/g) !== -1
    })
}

const getRows = (sheet) => {
    return readXlsxFile(base, { sheet: sheet.name });
}

const formatedData = (rows, sheetName) => {
    var agrupador = '';
    var newRows = [];

    for (const i in rows) {
        const row = rows[i]

        // header
        if (i == 0) {
            continue;
        }

        var agrupadores = ['Dívidas de implantações passadas', 'Novas implantações com receita associada (nova RRA)', 'Novas implantações sem receita associada (nova RRA)', 'TOTAL']

        if (agrupadores.indexOf(rows[i][0]) !== -1) {
            agrupador = rows[i][0];

        } else {
            // Fixa o agrupador na linha
            row.push(removeAcento(agrupador));
            row.push(sheetName);

            // remove os acentos da linhas
            row[0] = removeAcento(row[0]);

            if(sheetName !==  'Dezembro.20'){
                newRows.push([
                   row[0],         // 'Equipe de Implantacao',
                   row[1],         // 'Backlog inicial (saldo em horas)',
                   row[2],         // 'Horas executadas no mes',
                   row[3],         // 'Horas adicionadas no mes',
                   null,           // 'Horas de vendas no mês',
                   row[4],         // 'Backlog final (saldo em horas)',
                   row[5],         // 'Agrupador',
                   row[6],         // 'data'
                ])
            }else{
                newRows.push([
                    row[0],         // 'Equipe de Implantacao',
                    row[1],         // 'Backlog inicial (saldo em horas)',
                    row[2],         // 'Horas executadas no mes',
                    row[3],         // 'Horas adicionadas no mes',
                    row[4],         // 'Horas de vendas no mês',
                    row[5],         // 'Backlog final (saldo em horas)',
                    row[6],         // 'Agrupador',
                    row[7],         // 'data'
                 ])

            }

        }
    }

    return newRows;

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

const saveToCSV = async (arr) => {
    const csvFromArrayOfArrays = convertArrayToCSV(arr, {
        header,
        separator: ';'
    });

    fs = require('fs');
    fs.writeFileSync('./export/months.csv', csvFromArrayOfArrays,{encoding: 'utf8'})

   console.log("foi months");
}

// Executa
exports.execute = execute;
