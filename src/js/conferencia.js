const pProcess = document.getElementById('pProcessando')
const loader = document.getElementById('loader')

const inputGreenVoice = document.getElementById('fileGreenVoiceInput');
const inputAkiva = document.getElementById('fileAkivaInput');

const labelGreenVoice = document.getElementById('labelGreenVoiceFileInput')
const labelAkiva = document.getElementById('labelAkivaFileInput')

let filesGreenVoiceProcessed = 0
let filesAkivaProcessed = 0

let combinedGreenVoiceData = []
let combinedAkivaData = []



inputGreenVoice.addEventListener('change', (event) => {
    loader.setAttribute('style', 'display: block')
    pProcess.setAttribute('style', 'display: block')


    if (event.target.files.length > 1) {
        labelGreenVoice.textContent = `${event.target.files.length} arquivos`
    } else if (event.target.files.length == 1) {
        labelGreenVoice.textContent = `${event.target.files.length} arquivo`
    }

    const filesGreenVoice = inputGreenVoice.files;

    Array.from(filesGreenVoice).forEach(file => {
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
            combinedGreenVoiceData = combinedGreenVoiceData.concat(csvData);

            objGreenVoice = groupByIndices(combinedGreenVoiceData, 4);
        };

        reader.readAsArrayBuffer(file);
    });

    pProcess.setAttribute('style', 'display: none')
    loader.setAttribute('style', 'display: none')
})

inputAkiva.addEventListener('change', (event) => {
    if (event.target.files.length > 1) {
        labelAkiva.textContent = `${event.target.files.length} arquivos`
    } else if (event.target.files.length == 1) {
        labelAkiva.textContent = `${event.target.files.length} arquivo`
    }
})



document.getElementById('processBtn').addEventListener('click', function () {
    loader.setAttribute('style', 'display: block')
    pProcess.setAttribute('style', 'display: block')

    const filesAkiva = inputAkiva.files;


    if (filesAkiva.length === 0) {
        alert('Por favor, selecione pelo menos um arquivo.');
        return;
    }

    Array.from(filesAkiva).forEach(file => {
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
            combinedAkivaData = combinedAkivaData.concat(csvData);
            filesAkivaProcessed++;

            // Se todos os arquivos foram processados, faça a manipulação dos dados
            if (filesAkivaProcessed === filesAkiva.length) {
                manipulateAkivaData(combinedAkivaData);
            }
        };

        reader.readAsArrayBuffer(file);
    });
});

function groupByIndices(array, index) {
    const groups = {};

    array.forEach(subArray => {
        const key = subArray[index];
        if (!groups[key]) {
            groups[key] = { quantidade: 1 };
        } else {
            groups[key].quantidade++;
        }
    });

    return groups;
}

function groupByIndicesTabulacao(array, index1, index2) {
    const groups = {};

    array.forEach(subArray => {
        const key = `${subArray[index1]}${subArray[index2]}`;
        if (!groups[key]) {
            groups[key] = { quantidade: 1 };
        } else {
            groups[key].quantidade++;
        }
    });

    return groups;
}


function cruzarObjetos(obj1, obj2) {
    const resultado = { ...obj1 }; // Cria uma cópia de obj1 para manipular

    Object.keys(obj1).forEach(chave1 => {
        const chaveSem55 = chave1.startsWith("55") ? chave1.slice(2) : chave1;
        if (obj2[chaveSem55]) {
            obj1[chave1].quantidade -= obj2[chaveSem55].quantidade;
            if (obj1[chave1].quantidade <= 0) {
                delete obj1[chave1];
            }
        }
    });

    return obj1;
}

let objGreenVoice = {}
let objAkiva = {}


function manipulateAkivaData(data) {
    console.log(data)

    objAkiva = groupByIndicesTabulacao(data, 6, 7);
    console.log(objAkiva);

    Object.keys(objGreenVoice).forEach(chave1 => {
        const chaveSem55 = chave1.startsWith("55") ? chave1.slice(2) : chave1;
        if (objAkiva[chaveSem55]) {
            objGreenVoice[chave1].quantidade -= objAkiva[chaveSem55].quantidade;
            if (objGreenVoice[chave1].quantidade <= 0) {
                delete objGreenVoice[chave1];
            }
        }
    });

    // const resultado = cruzarObjetos(objGreenVoice, objAkiva)
    // console.log(resultado)


    pProcess.setAttribute('style', 'display: none')
    loader.setAttribute('style', 'display: none')

    const arrayDeArrays = Object.entries(objGreenVoice).map(([chave, valor]) => [chave, valor.quantidade]);
    console.log(arrayDeArrays.length);



    // exportToCSV(arrayDeArrays, 'Conferência', 'dataConferencia.xlsx')
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