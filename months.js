const readXlsxFile = require('read-excel-file/node');
const { convertArrayToCSV } = require('convert-array-to-csv');

const base = "./base/Alocação e produtividade das equipes.xlsx";

var header = [
    'Equipe de Implantação',
    'Backlog inicial (saldo em horas)',
    'Horas executadas no mês',
    'Horas adicionadas no mês',
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
            row.push(agrupador);
            row.push(sheetName);

            newRows.push(row);
        }
    }

    return newRows;

}

const saveToCSV = async (arr) => {
    const csvFromArrayOfArrays = convertArrayToCSV(arr, {
        header,
        separator: ';'
    });

    fs = require('fs');
    fs.writeFileSync('./export/months.csv', csvFromArrayOfArrays,{encoding: 'ascii'})

   console.log("foi months");
}

// Executa
exports.execute = execute;
