const pProcess = document.getElementById('pProcessando')
const loader = document.getElementById('loader')

const inputMovel = document.getElementById('fileMovelInput');
const inputFixa = document.getElementById('fileFixaInput');

const labelMovel = document.getElementById('labelMovelFileInput')
const labelFixa = document.getElementById('labelFixaFileInput')

let dataExport = []


let filesMovelProcessed = 0
let filesFixaProcessed = 0


let combinedMovelData = []
let combinedFixaData = []



inputMovel.addEventListener('change', (event) => {
    if (event.target.files.length > 1) {
        labelMovel.textContent = `${event.target.files.length} arquivos`
    } else if (event.target.files.length == 1) {
        labelMovel.textContent = `${event.target.files.length} arquivo`
    }
})

inputFixa.addEventListener('change', (event) => {
    if (event.target.files.length > 1) {
        labelFixa.textContent = `${event.target.files.length} arquivos`
    } else if (event.target.files.length == 1) {
        labelFixa.textContent = `${event.target.files.length} arquivo`
    }
})



document.getElementById('processBtn').addEventListener('click', function () {
    loader.setAttribute('style', 'display: block')
    pProcess.setAttribute('style', 'display: block')

    const filesMovel = inputMovel.files;
    const filesFixa = inputFixa.files;



    if (filesMovel.length === 0 || filesFixa.length === 0) {
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
                Array.from(filesFixa).forEach(file => {
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
                        combinedFixaData = combinedFixaData.concat(csvData);
                        filesFixaProcessed++;

                        // Se todos os arquivos foram processados, faça a manipulação dos dados
                        if (filesFixaProcessed === filesFixa.length) {
                            manipulateMovelData(combinedMovelData);
                        }
                    };

                    reader.readAsArrayBuffer(file);
                });
            }
        };

        reader.readAsArrayBuffer(file);
    });
});


function manipulateMovelData(data) {
    let vendasMovel = 0
    data.forEach(e => {
        if (e[10] == 'VENDA') {
            vendasMovel++
        }
    })
    dataExport.push({
        Categoria: 'Movel',
        Aceites: vendasMovel,
        Emissoes: '',
        Quebra: '',
        COD: 'COD',
        PercentCOD: '',
        GigasRedesSociais: 'Gigas Redes Sociais'
    })

    pProcess.setAttribute('style', 'display: none')
    loader.setAttribute('style', 'display: none')

    manipulateFixaData(combinedFixaData);
}



function manipulateFixaData(data) {
    let vendasFixa = 0

    data.forEach(e => {
        if (e[10] == 'VENDA') {
            vendasFixa++
        }
    })
    dataExport.push({
        Categoria: 'Fixa',
        Aceites: vendasFixa,
        Emissoes: '',
        Quebra: '',
        COD: 'COD',
        PercentCOD: '',
        GigasRedesSociais: 'Gigas Redes Sociais'
    })


    pProcess.setAttribute('style', 'display: none')
    loader.setAttribute('style', 'display: none')


    exportToCSV(dataExport, 'dataFichaDeVendas.xlsx')
}



function exportToCSV(data, filename) {
    // Cria uma nova worksheet a partir dos dados filtrados
    const worksheet1 = XLSX.utils.json_to_sheet(data);

    // Cria um novo workbook e adiciona a worksheet a ele
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet1, 'Ficha de Vendas');

    // Exporta o workbook como um arquivo XLSX
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx' });
}