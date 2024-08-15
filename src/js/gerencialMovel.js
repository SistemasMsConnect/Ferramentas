const pProcess = document.getElementById('pProcessando')
const loader = document.getElementById('loader')

let qnt = 0
let media = 0
let tempoMedio = 0
let dataExport = []

let qntCall = 0
let sumCall = 0
let mediaCall = 0


let filesProcessed = 0



document.getElementById('fileCallInput').addEventListener('change', (event) => {
    loader.setAttribute('style', 'display: block')
    pProcess.setAttribute('style', 'display: block')

    const file = event.target.files[0]

    const reader = new FileReader();

    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Assume que o CSV tem apenas uma folha (sheet)
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Converte o sheet para JSON
        const csvCallData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        manipulateCallData(csvCallData)

        pProcess.setAttribute('style', 'display: none')
        loader.setAttribute('style', 'display: none')
    };

    reader.readAsArrayBuffer(file);
})



function manipulateCallData(data) {
    data.forEach(e => {
        if (e[10] != 'Tempo_total') {
            const value = parseFloat(e[10]);

            if (!isNaN(value)) {
                sumCall += value;
                qntCall++
            }

        }

        return sumCall
    })

    mediaCall = sumCall / qntCall
}












document.getElementById('processBtn').addEventListener('click', function () {
    loader.setAttribute('style', 'display: block')
    pProcess.setAttribute('style', 'display: block')

    const input = document.getElementById('fileInput');
    const files = input.files;
    let combinedData = [];


    if (files.length === 0) {
        alert('Por favor, selecione pelo menos um arquivo CSV.');
        return;
    }

    Array.from(files).forEach(file => {
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
            combinedData = combinedData.concat(csvData);
            filesProcessed++;

            // Se todos os arquivos foram processados, faça a manipulação dos dados
            if (filesProcessed === files.length) {
                manipulateData(combinedData);
            }
        };

        reader.readAsArrayBuffer(file);
    });
});


function manipulateData(data) {
    const sum = sumColumn(data, 51);

    let naoAtende = 0
    let ocupado = 0
    let amd = 0
    let drop = 0

    let ligacaoInterrompida = 0
    let linhaMuda = 0
    let mensagemOperadora = 0
    let semCobertura = 0
    let clienteInsatisfeitoComAVivo = 0
    let semCondicoesFinanceiras = 0
    let recebeuOfertaDaConcorrencia = 0
    let vivoControle5GbIV = 0
    let vaiPensar = 0
    let responsavelNaoEstava = 0
    let mensalidadeAlta = 0
    let utilizaPoucoOCelular = 0
    let naoQuerCompromissoMensal = 0
    let clientePj = 0
    let clienteAbordadoRecentemente = 0
    let clienteMenorDeIdade = 0
    let falecimentoDoAssinante = 0

    data.forEach(e => {
        if (String(e[8]) == 'Falha da operadora') {
            naoAtende++
        } else if (String(e[8]).includes('Ocupado')) {
            ocupado++
        } else if (String(e[8]) == 'AMD') {
            amd++
        } else if (String(e[8]) == 'Drop') {
            drop++
        }

        if (String(e[10]).includes('INTERROMPIDA')) {
            ligacaoInterrompida++
        } else if (String(e[10]).includes('MUDA')) {
            linhaMuda++
        } else if (String(e[10]).includes('MENSAGEM OPERADORA')) {
            mensagemOperadora++
        } else if (String(e[10]).includes('SEM COBERTURA DISPONIBILIDADE')) {
            semCobertura++
        } else if (String(e[10]) == 'CLIENTE INSATISFEITO COM A VIVO') {
            clienteInsatisfeitoComAVivo++
        } else if (String(e[10]) == 'ACHA CARO SEM CONDICOES FINANCEIRAS') {
            semCondicoesFinanceiras++
        } else if (String(e[10]).includes('CONCORR')) {
            recebeuOfertaDaConcorrencia++
        } else if (String(e[10]) == 'VENDA') {
            vivoControle5GbIV++
        } else if (String(e[10]).includes('VAI PENSAR')) {
            vaiPensar++
        } else if (String(e[10]).includes('RESPONSAVEL')) {
            responsavelNaoEstava++
        } else if (String(e[10]).includes('MENSALIDADE ALTA')) {
            mensalidadeAlta++
        } else if (String(e[10]).includes('UTILIZA POUCO O CELULAR')) {
            utilizaPoucoOCelular++
        } else if (String(e[10]).includes('NAO QUER COMPROMISSO MENSAL')) {
            naoQuerCompromissoMensal++
        } else if (String(e[10]).includes('CLIENTE PJ')) {
            clientePj++
        } else if ( String(e[10]).includes('CLIENTE ABORDADO RECENTEMENTE')) {
            clienteAbordadoRecentemente++
        } else if (String(e[10]).includes('MENOR DE 18 ANOS')) {
            clienteMenorDeIdade++
        } else if (String(e[10]).includes('FALECIMENTO DO ASSINANTE')) {
            falecimentoDoAssinante++
        }
    })

    mediaTempo(sum, qnt)

    // console.log(data)

    dataExport.push({
        paPlanejadas: '-',
        paLogadas: '-',
        ocupacao: '',
        TMO: mediaCall,
        TMV: tempoMedio,
        _1: '',
        _2: '',
        _3: '',
        nomesDisponibilizados: '',
        _4: '',
        disparadasTotal: data.length - filesProcessed,
        _5: '',
        _6: '',
        _7: '',
        ligacoesAbandonadas: '-',
        _8: '',
        _9: '',
        drop: drop,
        _11: '',
        caixaPostal: '-',
        chamadaRejeitada: '-',
        falhaNaChamada: '-',
        ligacaoInterrompida: ligacaoInterrompida,
        linhaMuda: linhaMuda,
        mensagemOperadora: mensagemOperadora,
        naoAtende: naoAtende,
        ocupado: ocupado,
        amd: amd,
        _12: '',
        _13: '',
        _14: '',
        _15: '',
        _16: '',
        vaiPensar: vaiPensar,
        responsavelNaoEstava: responsavelNaoEstava,
        coringa1: '-',
        coringa2: '-',
        coringa3: '-',
        coringa4: '-',
        coringa5: '-',
        _17: '',
        _18: '',
        vivoControle4GbV: '-',
        vivoControle4GbIIIPln: '-',
        vivoControle5GbIV: vivoControle5GbIV,
        vivoControle5GbIPln: '-',
        vivoControle6GbIV: '-',
        vivoControle6GbIIPln: '-',
        vivoControle8GbIII: '-',
        vivoControle8GbIIIPln: '-',
        vivoControle10GbII: '-',
        vivoCtrlNovo10Gb_: '-',
        controleEducacao10Gb: '-',
        controleVantagens10Gb: '-',
        controleMusica10Gb: '-',
        controleSaude10Gb: '-',
        controleNetflix10Gb: '-',
        vivoControle10GbPlnI: '-',
        vivoControle12GbII: '-',
        vivoControle12GbPln: '-',
        controleEntretenimento14Gb: '-',
        coringa6: '-',
        coringa7: '-',
        coringa8: '-',
        coringa9: '-',
        coringa10: '-',
        coringa11: '-',
        coringa12: '-',
        coringa13: '-',
        coringa14: '-',
        coringa15: '-',
        coringa16: '-',
        coringa17: '-',
        coringa18: '-',
        coringa19: '-',
        _19: '',
        clienteInsatisfeitoComAVivo: clienteInsatisfeitoComAVivo,
        jaFoiControleENaoQuerVoltar: '-',
        temUmPosPagoDeOutraOperadora: '-',
        naoDesejaPassarPorAnaliseDeCredito: '-',
        naoQuerFazerMigracaoPorTelefone: '-',
        mensalidadeAlta: mensalidadeAlta,
        utilizaPoucoOCelular: utilizaPoucoOCelular,
        naoQuerCompromissoMensal: naoQuerCompromissoMensal,
        osPlanosOfertadosNaoAtendem: '-',
        semCondicoesFinanceiras: semCondicoesFinanceiras,
        desejaMaiorPacoteDeDados: '-',
        recebeuOfertaDaConcorrencia: recebeuOfertaDaConcorrencia,
        resistenteEmOuvirProposta: '-',
        clienteFidelizadoNaConcorrencia: '-',
        solicitadoPortabilidadeRecenteParaVivo: '-',
        clienteSolicitouPortabilidadeRecenteParaOutraOperadora: '-',
        clientePortadoParaOutraOperadoraAMenosDe90Dias: '-',
        _20: '',
        clientePj: clientePj,
        jaPossuiLinhaControlePosPagoVivo: '-',
        clienteAbordadoRecentemente: clienteAbordadoRecentemente,
        clienteMenorDeIdade: clienteMenorDeIdade,
        falecimentoDoAssinante: falecimentoDoAssinante,
        semCobertura: semCobertura,
        restricaoInterna: '-',
        restricaoExterna: '-',
        _21: '',
        _22: '',
        _23: '',
        _24: '',
        _25: '',
        _26: '',
        _27: '',
        _28: '',
        _29: '',
        _30: '',
        _31: '',
        _32: '',
        _33: '',
        _34: '',
        _35: '',
        _36: '',
        _37: '',
        _38: '',
        vivoControle4GbV2: '-',
        vivoControle4GbIIIPln2: '-',
        vivoControle5GbIV2: '',
        vivoControle5GbIPln2: '-',
        vivoControle6GbIV2: '-',
        vivoControle6GbIIPln2: '-',
        vivoControle8GbIII2: '-',
        vivoControle8GbIIIPln2: '-',
        vivoControle10GbII2: '-',
        vivoCtrlNovo10Gb_2: '-',
        controleEducacao10Gb2: '-',
        controleVantagens10Gb2: '-',
        controleMusica10Gb2: '-',
        controleSaude10Gb2: '-',
        controleNetflix10Gb2: '-',
        vivoControle10GbPlnI2: '-',
        vivoControle12GbII2: '-',
        vivoControle12GbPln2: '-',
        controleEntretenimento14Gb2: '-',
    })

    pProcess.setAttribute('style', 'display: none')
    loader.setAttribute('style', 'display: none')

    exportToCSV(dataExport, 'dataMovel.csv');
}


function mediaTempo(tempo, qntItens) {
    media = tempo / qntItens
    tempoMedio = media / 86400
}

function sumColumn(data, columnIndex) {
    let sum = 0

    data.forEach(row => {

        const value = parseFloat(row[columnIndex]);

        // Verifica se o valor é um número antes de somar
        if (!isNaN(value) && row[10] == 'venda' || row[10] == 'VENDA' || row[10] == 'Venda') {
            sum += value;
            qnt++
        }
    });

    return sum
}

function exportToCSV(data, filename) {
    // Cria uma nova worksheet a partir dos dados filtrados
    const worksheet = XLSX.utils.json_to_sheet(data);

    // Cria um novo workbook e adiciona a worksheet a ele
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Exporta o workbook como um arquivo CSV
    XLSX.writeFile(workbook, filename, { bookType: 'csv' });
}