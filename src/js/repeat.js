const pProcess = document.getElementById('pProcessando')
const loader = document.getElementById('loader')

const inputMovel = document.getElementById('fileMovelInput')
const inputDate = document.getElementById('dateInput')

const labelMovel = document.getElementById('labelMovelFileInput')

let dataExport = {}


let filesMovelProcessed = 0


let combinedMovelData = []



inputMovel.addEventListener('change', (event) => {
    if (event.target.files.length > 1) {
        labelMovel.textContent = `${event.target.files.length} arquivos`
    } else if (event.target.files.length == 1) {
        labelMovel.textContent = `${event.target.files.length} arquivo`
    }
})



document.getElementById('processBtn').addEventListener('click', function () {
    loader.setAttribute('style', 'display: block')
    pProcess.setAttribute('style', 'display: block')

    const filesMovel = inputMovel.files;


    if (filesMovel.length === 0) {
        alert('Por favor, selecione pelo menos um arquivo CSV.');
        return;
    }

    Array.from(filesMovel).forEach(file => {
        const reader = new FileReader();

        reader.onload = function (event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Assume que o CSV tem apenas uma folha (sheet)
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Converte o sheet para JSON
            const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Adiciona os dados ao array combinado
            combinedMovelData = combinedMovelData.concat(csvData);
            filesMovelProcessed++;

            // Se todos os arquivos foram processados, faça a manipulação dos dados
            if (filesMovelProcessed === filesMovel.length) {
                manipulateMovelData(combinedMovelData);
            }
        };

        reader.readAsArrayBuffer(file);
    });
});


function manipulateMovelData(data) {
    console.log(data)

    data.forEach(e => {
        const key = `${e[5]}${e[6]}`

        if (!dataExport[key]) {
            dataExport[key] = {
                total: 0,           // Contador total para a chave
                index8: {},         // Contagem dos valores do índice 8
                index10: {}         // Contagem dos valores do índice 10
            };
        }

        // Incrementa o contador total para a combinação
        dataExport[key].total += 1;

        // Incrementa a contagem para o valor no índice 8
        const value8 = e[8];
        dataExport[key].index8[value8] = (dataExport[key].index8[value8] || 0) + 1;

        // Incrementa a contagem para o valor no índice 10
        const value10 = e[10];
        dataExport[key].index10[value10] = (dataExport[key].index10[value10] || 0) + 1;
    });

    console.log(dataExport)

    exportCountMapToCSV(dataExport, 'Contagens', `Contagens_${inputDate.value}.xlsx`)

    pProcess.setAttribute('style', 'display: none')
    loader.setAttribute('style', 'display: none')
}



function exportCountMapToCSV(countMap, sheetname, filename) {
    // Transforma o objeto countMap em um array de objetos para exportação
    const data = Object.keys(countMap).map(key => {
        const entry = countMap[key];

        // Cria um objeto com as informações do countMap, incluindo contagens de index8 e index10
        const result = {
            combination: key,
            total: entry.total
        };

        // Adiciona contagens de index8 ao objeto result
        Object.keys(entry.index8).forEach(value8 => {
            result[`ISDN_${value8}`] = entry.index8[value8];
        });

        // Adiciona contagens de index10 ao objeto result
        Object.keys(entry.index10).forEach(value10 => {
            result[`Classificação_${value10}`] = entry.index10[value10];
        });

        return result;
    });

    // Cria uma nova worksheet a partir dos dados de countMap
    const worksheet = XLSX.utils.json_to_sheet(data);

    // Cria um novo workbook e adiciona a worksheet a ele
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetname);

    // Exporta o workbook como um arquivo XLSX
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx' });
}








function exportToCSV(data, sheetname, filename) {
    // Cria uma nova worksheet a partir dos dados filtrados
    const worksheet1 = XLSX.utils.json_to_sheet(data);

    // Cria um novo workbook e adiciona a worksheet a ele
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet1, `${sheetname}`);

    // Exporta o workbook como um arquivo XLSX
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx' });
}