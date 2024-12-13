const pProcess = document.getElementById('pProcessando')
const loader = document.getElementById('loader')

const inputInput = document.getElementById('fileInputInput');

const labelInput = document.getElementById('labelInputFileInput')

let dataInputExport = []


let filesInputProcessed = 0


let combinedInputData = []



inputInput.addEventListener('change', (event) => {
    labelInput.textContent = event.target.files[0].name
})



document.getElementById('processBtn').addEventListener('click', function () {
    loader.setAttribute('style', 'display: block')
    pProcess.setAttribute('style', 'display: block')

    const filesInput = inputInput.files;

    if (filesInput.length === 0) {
        alert('Por favor, selecione pelo menos um arquivo.');
        return;
    }

    Array.from(filesInput).forEach(file => {
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
            combinedInputData = combinedInputData.concat(csvData);
            manipulateInputData(combinedInputData);
        };

        reader.readAsArrayBuffer(file);
    });
})



function manipulateInputData(data) {
    console.log(data)

    data.forEach(e => {
        let regiao = ''

        if (['11', '12', '13'].includes(String(e[7]).slice(0, 2))) {
            regiao = 'SP1'
        } else if (['14', '15', '16', '17', '18', '19'].includes(String(e[7]).slice(0, 2))) {
            regiao = 'SPI'
        }

        dataInputExport.push({
            iD: e[0],
            Fila: e[29],
            Regiao: regiao,
            DDD: String(e[7]).slice(0, 2),
            Telefone: e[7],
            Agente: e[37],
            Status: e[48]
        })
    })

    exportToCSV(dataInputExport, 'dataInput.xlsx');

    loader.setAttribute('style', 'display: none')
    pProcess.setAttribute('style', 'display: none')
}



function exportToCSV(dataInput, filename) {
    // Cria uma nova worksheet a partir dos dados filtrados
    const worksheet1 = XLSX.utils.json_to_sheet(dataInput);

    // Cria um novo workbook e adiciona a worksheet a ele
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet1, 'Input');

    // Exporta o workbook como um arquivo XLSX
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx' });
}