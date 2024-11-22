const loader = document.getElementById('loader')
const pProcess = document.getElementById('pProcessando')

const labelMovel = document.getElementById('labelMovelFileInput')
const labelRecomendacao = document.getElementById('labelRecomendacaoFileInput')
const labelSuspensao = document.getElementById('labelSuspensaoFileInput')

const inputMovel = document.getElementById('movelFileInput')
const inputRecomendacao = document.getElementById('recomendacaoFileInput')
const inputSuspensao = document.getElementById('suspensaoFileInput')

let filesMovelProcessed = 0
let filesRecomendacaoProcessed = 0
let filesSuspensaoProcessed = 0

let combinedMovelData = []
let combinedRecomendacaoData = []
let combinedSuspensaoData = []

let resultadoExport = []

inputMovel.addEventListener('change', (event) => {
    if (event.target.files.length > 1) {
        labelMovel.textContent = `${event.target.files.length} arquivos`
    } else if (event.target.files.length == 1) {
        labelMovel.textContent = `${event.target.files.length} arquivo`
    }
})

inputRecomendacao.addEventListener('change', (event) => {
    if (event.target.files.length > 1) {
        labelRecomendacao.textContent = `${event.target.files.length} arquivos`
    } else if (event.target.files.length == 1) {
        labelRecomendacao.textContent = `${event.target.files.length} arquivo`
    }
})

inputSuspensao.addEventListener('change', (event) => {
    if (event.target.files.length > 1) {
        labelSuspensao.textContent = `${event.target.files.length} arquivos`
    } else if (event.target.files.length == 1) {
        labelSuspensao.textContent = `${event.target.files.length} arquivo`
    }
})


document.getElementById('processBtn').addEventListener('click', function () {
    loader.setAttribute('style', 'display: block')
    pProcess.setAttribute('style', 'display: block')

    const filesMovel = inputMovel.files;
    const filesRecomendacao = inputRecomendacao.files;
    const filesSuspensao = inputSuspensao.files;


    if (filesMovel.length === 0 || filesRecomendacao.length === 0 || filesSuspensao.length === 0) {
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
                Array.from(filesRecomendacao).forEach(file => {
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
                        combinedRecomendacaoData = combinedRecomendacaoData.concat(csvData);
                        filesRecomendacaoProcessed++;

                        // Se todos os arquivos foram processados, faça a manipulação dos dados
                        if (filesRecomendacaoProcessed === filesRecomendacao.length) {
                            Array.from(filesSuspensao).forEach(file => {
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
                                    combinedSuspensaoData = combinedSuspensaoData.concat(csvData);
                                    filesSuspensaoProcessed++;

                                    // Se todos os arquivos foram processados, faça a manipulação dos dados
                                    if (filesSuspensaoProcessed === filesSuspensao.length) {
                                        manipulateSuspensaoData(combinedSuspensaoData)
                                    }
                                };

                                reader.readAsArrayBuffer(file);
                            });
                        }
                    };

                    reader.readAsArrayBuffer(file);
                });
            }
        };

        reader.readAsArrayBuffer(file);
    });
});


function manipulateSuspensaoData(data) {
    console.log(data)

    manipulateRecomendacaoData(combinedRecomendacaoData)
}

function manipulateRecomendacaoData(data) {
    console.log(data)
    data.forEach(e => {
        if (e[10] != undefined) {
            e[10] = e[10].trim()
        }
        if (e[11] != undefined) {
            e[11] = e[11].trim()
        }
    })

    manipulateMovelData(combinedMovelData)
}

function manipulateMovelData(data) {
    console.log(data)

    // const headers = data[0]
    // const dataObject = data.slice(1).map(linha => Object.fromEntries(headers.map((key, index) => [key, linha[index]])))

    // dataObject.forEach(subObj => {
    //     if (subObj['M'] == 'M') {
    //         subObj['M'] = 'mInutilizado'
    //     }

    //     if (subObj['M_RECOMENDACAO'] == 'M_RECOMENDACAO') {
    //         subObj['M_RECOMENDACAO'] = 'M'
    //     } else if (subObj['M_RECOMENDACAO'] < 17) {
    //         subObj['FIDELIZADO'] = 'Fidelizado'
    //     } else if (subObj['M_RECOMENDACAO']) {
    //         subObj['FIDELIZADO'] = 'Não Fidelizado'
    //     }

    //     let indexS = combinedSuspensaoData.findIndex(el => el[11] == subObj['NR_TELEFONE'])
    // })

    data.forEach(subAr => {
        if (subAr[12] == 'M_RECOMENDACAO') {
            subAr[12] = 'M'
        } else if (subAr[12] < 17) {
            subAr[16] = 'Fidelizado'
        } else if (subAr[12] > 16) {
            subAr[16] = 'Não Fidelizado'
        }

        let indexS = combinedSuspensaoData.findIndex(el => el[11] == subAr[7])
        let indexR = combinedRecomendacaoData.findIndex(el => el[5] == subAr[7])

        let linhaSuspensa = ''
        let linhaSuspensaStatus = ''

        if (indexS != -1) {
            linhaSuspensaStatus = combinedSuspensaoData[indexS][9]
            linhaSuspensa = 'Sim'
        } else {
            linhaSuspensaStatus = '0'
            linhaSuspensa = 'Não'
        }

        if (indexR != -1) {
            subAr[19] = combinedRecomendacaoData[indexR][10]
            subAr[20] = combinedRecomendacaoData[indexR][11]
        } else {
            subAr[19] = '0'
            subAr[20] = '0'
        }

        resultadoExport.push({
            CODCLI: subAr[4],
            CNPJ_CLIENTE: subAr[5],
            NR_TELEFONE: subAr[7],
            UF: subAr[8],
            PLANO: subAr[9],
            M: subAr[12],
            FIDELIZADO: subAr[16],
            DT_PRIMEITA_ATIVACAO: subAr[17],
            PLANO_CONTA: subAr[19],
            PLANO_LINHA: subAr[20],
            LINHA_SUSPENSA: linhaSuspensa,
            LINHA_SUSPENSA_STATUS: linhaSuspensaStatus,
            FAT_MEDIO_03_MESES: subAr[10],
            QT_DIAS_TRAF_DADOS: subAr[13],
            QT_MB_TRAF_DADOS: subAr[15],
        })
    })

    exportToXLSX(resultadoExport.slice(1), 'movel_Data.xlsx')


    loader.setAttribute('style', 'display: none')
    pProcess.setAttribute('style', 'display: none')
}



function exportToXLSX(data, filename) {
    // Cria uma nova worksheet a partir dos dados filtrados
    const worksheet1 = XLSX.utils.json_to_sheet(data);

    // Cria um novo workbook e adiciona a worksheet a ele
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet1, 'Movel');

    // Exporta o workbook como um arquivo XLSX
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx' });
}