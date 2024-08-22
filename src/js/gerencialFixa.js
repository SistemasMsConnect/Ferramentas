const pProcess = document.getElementById('pProcessando')
const loader = document.getElementById('loader')

let qnt = 0
let media = 0
let tempoMedio = 0
let dataExport = []

let qntCall = 0
let sumCall = 0
let mediaCall = 0
let qntCallVenda = 0
let sumCallVenda = 0
let mediaCallVenda = 0


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
    console.log(data)
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

    data.forEach(e => {
        if (e[5] == 'VENDA') {
            if (e[10] != 'Tempo_total') {
                const value = parseFloat(e[10]);

                if (!isNaN(value)) {
                    sumCallVenda += value;
                    qntCallVenda++
                }

            }
        }
        return sumCallVenda
    })

    mediaCall = sumCall / qntCall
    mediaCallVenda = sumCallVenda / qntCallVenda
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

    let ligacaoInterrompida = 0
    let linhaMuda = 0
    let mensagemOperadora = 0
    let semCobertura = 0
    let enderecoNaoPertenceAoContato = 0
    let naoUtilizaraOProduto = 0
    let naoVeBeneficiosNaOferta = 0
    let clienteInsatisfeitoComAVivo = 0
    let jaPossuiProduto = 0
    let achaCaro = 0
    let possuiConcorrencia = 0
    let naoQuerMaisReceberLigacoesDaVivo = 0
    let portabilidadeEmProcesso = 0
    let telefoneNaoPertenceAoContato = 0
    let vivoInternetFibra = 0

    data.forEach(e => {
        if (String(e[8]).includes('TimeOut')) {
            naoAtende++
        } else if (String(e[8]).includes('Ocupado')) {
            ocupado++
        } else if (String(e[8]) == 'AMD') {
            amd++
        }

        if (String(e[10]).includes('INTERROMPIDA') && !String(e[10]).includes('SEM CLIENTE')) {
            ligacaoInterrompida++
        } else if (String(e[10]).includes('MUDA')) {
            linhaMuda++
        } else if (String(e[10]).includes('MENSAGEM OPERADORA')) {
            mensagemOperadora++
        } else if (String(e[10]).includes('SEM COBERTURA DISPONIBILIDADE')) {
            semCobertura++
        } else if (String(e[10]).includes('ENDERE') && !String(e[10]).includes('DISPONIBILIZADA')) {
            enderecoNaoPertenceAoContato++
        } else if (String(e[10]).includes('O PRODUTO')) {
            naoUtilizaraOProduto++
        } else if (String(e[10]) == 'NAO VE BENEFICIOS NA OFERTA') {
            naoVeBeneficiosNaOferta++
        } else if (String(e[10]) == 'CLIENTE INSATISFEITO COM A VIVO') {
            clienteInsatisfeitoComAVivo++
        } else if (String(e[10]) == 'JA POSSUI PRODUTO TEM OS PENDENTE') {
            jaPossuiProduto++
        } else if (String(e[10]) == 'ACHA CARO SEM CONDICOES FINANCEIRAS') {
            achaCaro++
        } else if (String(e[10]).includes('CONCORR')) {
            possuiConcorrencia++
        } else if (String(e[10]) == 'NAO QUER MAIS RECEBER LIGACOES DA VIVO') {
            naoQuerMaisReceberLigacoesDaVivo++
        } else if (String(e[10]) == 'PORTABILIDADE EM PROCESSO') {
            portabilidadeEmProcesso++
        } else if (String(e[10]) == 'TELEFONE NAO PERTENCE AO CONTATO') {
            telefoneNaoPertenceAoContato++
        } else if (String(e[10]) == 'VENDA') {
            vivoInternetFibra++
        }
    })

    mediaTempo(sum, qnt)

    // console.log(data)

    dataExport.push({
        paPlanejadas: '-',
        paLogadas: '-',
        ocupacao: '',
        TMO: mediaCall,
        TMV: mediaCallVenda,
        _1: '',
        _2: '',
        _3: '',
        _4: '',
        cobre: '-',
        fwt: '-',
        voip: '-',
        _5: '',
        _6: '',
        _7: '',
        _8: '',
        _9: '',
        _10: '',
        _11: '',
        _12: '',
        _13: '',
        _14: '',
        _15: '',
        _16: '',
        _17: '',
        _18: '',
        _19: '',
        _20: '',
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
        contaProtegidaMapfre: '-',
        seguroResidencial: '-',
        contaDigital: '-',
        vivoProtege: '-',
        coringa1: '-',
        _35: '',
        vivoFixo: '-',
        vivoInternetFixa: '-',
        vivoInternetFibra1: '-',
        vivoTv: '-',
        vivoTvFibra: '-',
        _36: '',
        _37: '',
        _38: '',
        _39: '',
        _40: '',
        _41: '',
        _42: '',
        _43: '',
        _44: '',
        _45: '',
        _46: '',
        _47: '',
        _48: '',
        _49: '',
        _50: '',
        _51: '',
        _52: '',
        _53: '',
        _54: '',
        _55: '',
        _56: '',
        _57: '',
        _58: '',
        _59: '',
        _60: '',
        _61: '',
        _62: '',
        _63: '',
        _64: '',
        _65: '',
        _66: '',
        _67: '',
        _68: '',
        _69: '',
        _70: '',
        _71: '',
        _72: '',
        _73: '',
        _74: '',
        _75: '',
        _76: '',
        _77: '',
        _78: '',
        _79: '',
        _80: '',
        _81: '',
        _82: '',
        _83: '',
        _84: '',
        _85: '',
        _86: '',
        _87: '',
        _88: '',
        _89: '',
        _90: '',
        _91: '',
        _92: '',
        _93: '',
        _94: '',
        _95: '',
        _96: '',
        _97: '',
        _98: '',
        _99: '',
        _100: '',
        _101: '',
        _102: '',
        _103: '',
        _104: '',
        _105: '',
        _106: '',
        _107: '',
        _108: '',
        _109: '',
        _110: '',
        _111: '',
        _112: '',
        _113: '',
        _114: '',
        _115: '',
        _116: '',
        _117: '',
        mailingRecebido: '',
        mailingCarregado: '',
        mailingCarregadoRepique: '-',
        blacklist: '-',
        duplicidade: '-',
        outros: '-',
        nomesTrabalhados: data.length - filesProcessed,
        nomesVirgens: '',
        _118: '',
        _119: '',
        _120: '',
        _121: '',
        _122: '',
        _123: '',
        naoAtende: naoAtende,
        ligacaoInterrompida: ligacaoInterrompida,
        ocupado: ocupado,
        linhaMuda: linhaMuda,
        mensagemOperadora: mensagemOperadora,
        amd: amd,
        coringa2: '-',
        coringa3: '-',
        coringa4: '-',
        _124: '',
        _125: '',
        _126: '',
        _127: '',
        naoAtende1: '-',
        ligacaoInterrompida1: '-',
        ocupado1: '-',
        linhaMuda1: '-',
        mensagemOperadora1: '-',
        amd1: '-',
        coringa5: '-',
        coringa6: '-',
        coringa7: '-',
        _128: '',
        coringa8: '-',
        agendamento: '-',
        coringa9: '-',
        _129: '',
        _130: '',
        _131: '',
        _132: '',
        _133: '',
        _134: '',
        _135: '',
        _136: '',
        _137: '',
        _138: '',
        _139: '',
        criticaSistemicaSiebel1: '-',
        cpfInvalido1: '-',
        creditoNegado1: '-',
        semCobertura1: '-',
        enderecoNaoPertenceAoContato1: '-',
        coringa10: '-',
        coringa11: '-',
        coringa12: '-',
        coringa13: '-',
        coringa14: '-',
        coringa15: '-',
        coringa16: '-',
        _140: '',
        _141: '',
        criticaSistemicaSiebel2: '-',
        cpfInvalido2: '-',
        creditoNegado2: '-',
        semCobertura2: '-',
        enderecoNaoPertenceAoContato2: '-',
        coringa17: '-',
        coringa18: '-',
        coringa19: '-',
        coringa20: '-',
        coringa21: '-',
        coringa22: '-',
        coringa23: '-',
        _142: '',
        _143: '',
        criticaSistemicaSiebel3: '-',
        cpfInvalido3: '-',
        creditoNegado3: '-',
        semCobertura3: '-',
        enderecoNaoPertenceAoContato3: '-',
        coringa24: '-',
        coringa25: '-',
        coringa26: '-',
        coringa27: '-',
        coringa28: '-',
        coringa29: '-',
        coringa30: '-',
        _144: '',
        _145: '',
        criticaSistemicaSiebel: '-',
        cpfInvalido: '-',
        creditoNegado: '-',
        semCobertura: semCobertura,
        enderecoNaoPertenceAoContato: enderecoNaoPertenceAoContato,
        coringa31: '-',
        coringa32: '-',
        coringa33: '-',
        coringa34: '-',
        coringa35: '-',
        coringa36: '-',
        coringa37: '-',
        _146: '',
        _147: '',
        criticaSistemicaSiebel4: '-',
        cpfInvalido4: '-',
        creditoNegado4: '-',
        semCobertura4: '-',
        enderecoNaoPertenceAoContato4: '-',
        coringa38: '-',
        coringa39: '-',
        coringa40: '-',
        coringa41: '-',
        coringa42: '-',
        coringa43: '-',
        coringa44: '-',
        _148: '',
        _149: '',
        clienteAbordadoRecentemente: '-',
        clienteComMenosDe3MesesNaBase: '-',
        clientePj: '-',
        falecimentoDoAssinante: '-',
        menorDe18Anos: '-',
        naoPossuiOsDadosDoTitular: '-',
        RestricaoExterna: '-',
        RestricaoInterna: '-',
        coringa45: '-',
        coringa46: '-',
        coringa47: '-',
        coringa48: '-',
        _150: '',
        _151: '',
        _152: '',
        naoUtilizaraOProduto1: '-',
        naoVeBeneficiosNaOferta1: '-',
        seRecusaAFalar1: '-',
        clienteInsatisfeitoComAVivo1: '-',
        jaPossuiProduto1: '-',
        achaCaro1: '-',
        possuiConcorrencia1: '-',
        naoQuerMaisReceberLigacoesDaVivo1: '-',
        portabilidadeEmProcesso1: '-',
        telefoneNaoPertenceAoContato1: '-',
        coringa49: '-',
        coringa50: '-',
        coringa51: '-',
        coringa52: '-',
        _153: '',
        _154: '',
        naoUtilizaraOProduto2: '-',
        naoVeBeneficiosNaOferta2: '-',
        seRecusaAFalar2: '-',
        clienteInsatisfeitoComAVivo2: '-',
        jaPossuiProduto2: '-',
        achaCaro2: '-',
        possuiConcorrencia2: '-',
        naoQuerMaisReceberLigacoesDaVivo2: '-',
        portabilidadeEmProcesso2: '-',
        telefoneNaoPertenceAoContato2: '-',
        coringa53: '-',
        coringa54: '-',
        coringa55: '-',
        coringa56: '-',
        _155: '',
        _156: '',
        naoUtilizaraOProduto3: '-',
        naoVeBeneficiosNaOferta3: '-',
        seRecusaAFalar3: '-',
        clienteInsatisfeitoComAVivo3: '-',
        jaPossuiProduto3: '-',
        achaCaro3: '-',
        possuiConcorrencia3: '-',
        naoQuerMaisReceberLigacoesDaVivo3: '-',
        portabilidadeEmProcesso3: '-',
        telefoneNaoPertenceAoContato3: '-',
        coringa57: '-',
        coringa58: '-',
        coringa59: '-',
        coringa60: '-',
        _157: '',
        _158: '',
        naoUtilizaraOProduto4: '-',
        naoVeBeneficiosNaOferta4: '-',
        seRecusaAFalar4: '-',
        clienteInsatisfeitoComAVivo4: '-',
        jaPossuiProduto4: '-',
        achaCaro4: '-',
        possuiConcorrencia4: '-',
        naoQuerMaisReceberLigacoesDaVivo4: '-',
        portabilidadeEmProcesso4: '-',
        telefoneNaoPertenceAoContato4: '-',
        coringa61: '-',
        coringa62: '-',
        coringa63: '-',
        coringa64: '-',
        _159: '',
        _160: '',
        naoUtilizaraOProduto: naoUtilizaraOProduto,
        naoVeBeneficiosNaOferta: naoVeBeneficiosNaOferta,
        seRecusaAFalar: '-',
        clienteInsatisfeitoComAVivo: clienteInsatisfeitoComAVivo,
        jaPossuiProduto: jaPossuiProduto,
        achaCaro: achaCaro,
        possuiConcorrencia: possuiConcorrencia,
        naoQuerMaisReceberLigacoesDaVivo: naoQuerMaisReceberLigacoesDaVivo,
        portabilidadeEmProcesso: portabilidadeEmProcesso,
        telefoneNaoPertenceAoContato: telefoneNaoPertenceAoContato,
        coringa65: '-',
        coringa66: '-',
        coringa67: '-',
        coringa68: '-',
        _161: '',
        _162: '',
        naoUtilizaraOProduto5: '-',
        naoVeBeneficiosNaOferta5: '-',
        seRecusaAFalar5: '-',
        clienteInsatisfeitoComAVivo5: '-',
        jaPossuiProduto5: '-',
        achaCaro5: '-',
        possuiConcorrencia5: '-',
        naoQuerMaisReceberLigacoesDaVivo5: '-',
        portabilidadeEmProcesso5: '-',
        telefoneNaoPertenceAoContato5: '-',
        coringa69: '-',
        coringa70: '-',
        coringa71: '-',
        coringa72: '-',
        _163: '',
        _164: '',
        naoUtilizaraOProduto6: '-',
        naoVeBeneficiosNaOferta6: '-',
        seRecusaAFalar6: '-',
        clienteInsatisfeitoComAVivo6: '-',
        jaPossuiProduto6: '-',
        achaCaro6: '-',
        possuiConcorrencia6: '-',
        naoQuerMaisReceberLigacoesDaVivo6: '-',
        portabilidadeEmProcesso6: '-',
        telefoneNaoPertenceAoContato6: '-',
        coringa73: '-',
        coringa74: '-',
        coringa75: '-',
        coringa76: '-',
        _165: '',
        _166: '',
        insegurancaComExcedenteDefatura: '-',
        mensalidadeAlta: '-',
        naoQuerCompromissoMensal: '-',
        naoVeVantagemNaTrocaDePlano: '-',
        poucaFranquiaDeDados: '-',
        poucoBeneficioParaOutrasOperadoras: '-',
        querOfertaDeAparelho: '-',
        coringa77: '-',
        coringa78: '-',
        coringa79: '-',
        coringa80: '-',
        coringa81: '-',
        coringa82: '-',
        coringa83: '-',
        coringa84: '-',
        _167: '',
        _168: '',
        _169: '',
        vivoFixoIlimitadoBrasil: '-',
        vivoFixoIlimitadoLocal: '-',
        coringa85: '-',
        coringa86: '-',
        coringa87: '-',
        coringa88: '-',
        coringa89: '-',
        coringa90: '-',
        coringa91: '-',
        coringa92: '-',
        coringa93: '-',
        coringa94: '-',
        coringa95: '-',
        coringa96: '-',
        coringa97: '-',
        coringa98: '-',
        coringa99: '-',
        coringa100: '-',
        coringa101: '-',
        coringa102: '-',
        coringa103: '-',
        coringa104: '-',
        coringa105: '-',
        coringa106: '-',
        coringa107: '-',
        coringa108: '-',
        coringa109: '-',
        coringa110: '-',
        coringa111: '-',
        coringa112: '-',
        coringa113: '-',
        coringa114: '-',
        coringa115: '-',
        coringa116: '-',
        coringa117: '-',
        coringa118: '-',
        coringa119: '-',
        coringa120: '-',
        coringa121: '-',
        coringa122: '-',
        coringa123: '-',
        coringa124: '-',
        coringa125: '-',
        _170: '',
        vivoAssistenciaCasa: '-',
        coringa126: '-',
        coringa127: '-',
        coringa128: '-',
        coringa129: '-',
        coringa130: '-',
        coringa131: '-',
        coringa132: '-',
        coringa133: '-',
        _171: '',
        coringa134: '-',
        coringa135: '-',
        coringa136: '-',
        coringa137: '-',
        coringa138: '-',
        coringa139: '-',
        coringa140: '-',
        coringa141: '-',
        coringa142: '-',
        coringa143: '-',
        coringa144: '-',
        coringa145: '-',
        coringa146: '-',
        coringa147: '-',
        coringa148: '-',
        coringa149: '-',
        coringa150: '-',
        coringa151: '-',
        coringa152: '-',
        coringa153: '-',
        coringa154: '-',
        coringa155: '-',
        coringa156: '-',
        coringa157: '-',
        coringa158: '-',
        coringa159: '-',
        coringa160: '-',
        coringa161: '-',
        coringa162: '-',
        coringa163: '-',
        coringa164: '-',
        coringa165: '-',
        coringa166: '-',
        coringa167: '-',
        coringa168: '-',
        coringa169: '-',
        coringa170: '-',
        coringa171: '-',
        coringa172: '-',
        coringa173: '-',
        coringa174: '-',
        coringa175: '-',
        coringa176: '-',
        _172: '',
        coringa177: '-',
        coringa178: '-',
        coringa179: '-',
        coringa180: '-',
        coringa181: '-',
        coringa182: '-',
        coringa183: '-',
        coringa184: '-',
        coringa185: '-',
        coringa186: '-',
        coringa187: '-',
        coringa188: '-',
        coringa189: '-',
        coringa190: '-',
        coringa191: '-',
        coringa192: '-',
        coringa193: '-',
        coringa194: '-',
        coringa195: '-',
        coringa196: '-',
        coringa197: '-',
        coringa198: '-',
        coringa199: '-',
        coringa200: '-',
        coringa201: '-',
        coringa202: '-',
        coringa203: '-',
        coringa204: '-',
        coringa205: '-',
        coringa206: '-',
        coringa207: '-',
        coringa208: '-',
        coringa209: '-',
        coringa210: '-',
        coringa211: '-',
        coringa212: '-',
        coringa213: '-',
        coringa214: '-',
        coringa215: '-',
        coringa216: '-',
        coringa217: '-',
        coringa218: '-',
        coringa219: '-',
        _173: '',
        _1Mbps: '-',
        _2Mbps: '-',
        _4Mbps: '-',
        _8Mbps: '-',
        _10Mbps: '-',
        coringa220: '-',
        coringa221: '-',
        coringa222: '-',
        coringa223: '-',
        coringa224: '-',
        coringa225: '-',
        coringa226: '-',
        coringa227: '-',
        coringa228: '-',
        coringa229: '-',
        _174: '',
        coringa230: '-',
        coringa231: '-',
        coringa232: '-',
        coringa233: '-',
        coringa234: '-',
        coringa235: '-',
        coringa236: '-',
        coringa237: '-',
        coringa238: '-',
        coringa239: '-',
        coringa240: '-',
        coringa241: '-',
        coringa242: '-',
        coringa243: '-',
        coringa244: '-',
        _175: '',
        coringa245: '-',
        coringa246: '-',
        coringa247: '-',
        coringa248: '-',
        coringa249: '-',
        coringa250: '-',
        coringa251: '-',
        coringa252: '-',
        coringa253: '-',
        coringa254: '-',
        coringa255: '-',
        coringa256: '-',
        coringa257: '-',
        coringa258: '-',
        coringa259: '-',
        _176: '',
        coringa260: '-',
        coringa261: '-',
        coringa262: '-',
        coringa263: '-',
        coringa264: '-',
        coringa265: '-',
        coringa266: '-',
        coringa267: '-',
        coringa268: '-',
        coringa269: '-',
        coringa270: '-',
        coringa271: '-',
        coringa272: '-',
        coringa273: '-',
        coringa274: '-',
        _177: '',
        _5Mbps: '-',
        _15Mbps: '-',
        _25Mbps: '-',
        _50Mbps1: '-',
        coringa275: '-',
        coringa276: '-',
        coringa277: '-',
        coringa278: '-',
        coringa279: '-',
        coringa280: '-',
        coringa281: '-',
        coringa282: '-',
        coringa283: '-',
        coringa284: '-',
        coringa285: '-',
        _178: '',
        coringa286: '-',
        coringa287: '-',
        coringa288: '-',
        coringa289: '-',
        coringa290: '-',
        coringa291: '-',
        coringa292: '-',
        coringa293: '-',
        coringa294: '-',
        coringa295: '-',
        coringa296: '-',
        coringa297: '-',
        coringa298: '-',
        coringa299: '-',
        coringa300: '-',
        _179: '',
        _50Mbps2: '-',
        _70Mbps1: '-',
        _100Mbps1: '-',
        _200Mbps1: '-',
        _300Mbps1: '-',
        _600Mbps1: '-',
        vivoInternetFibra2: vivoInternetFibra,
        coringa301: '-',
        coringa302: '-',
        coringa303: '-',
        coringa304: '-',
        coringa305: '-',
        coringa306: '-',
        coringa307: '-',
        coringa308: '-',
        _180: '',
        coringa309: '-',
        coringa310: '-',
        coringa311: '-',
        coringa312: '-',
        coringa313: '-',
        coringa314: '-',
        coringa315: '-',
        coringa316: '-',
        coringa317: '-',
        coringa318: '-',
        coringa319: '-',
        coringa320: '-',
        coringa321: '-',
        coringa322: '-',
        coringa323: '-',
        _181: '',
        disneyPlus: '-',
        netflixPadrao: '-',
        netflixPremium: '-',
        globoplay: '-',
        amazon: '-',
        coringa324: '-',
        coringa325: '-',
        coringa326: '-',
        coringa327: '-',
        coringa328: '-',
        coringa329: '-',
        coringa330: '-',
        coringa331: '-',
        coringa332: '-',
        coringa333: '-',
        _182: '',
        _50Mbps3: '-',
        _70Mbps2: '-',
        _100Mbps2: '-',
        _200Mbps2: '-',
        _300Mbps2: '-',
        _600Mbps2: '-',
        coringa334: '-',
        coringa335: '-',
        coringa336: '-',
        coringa337: '-',
        coringa338: '-',
        coringa339: '-',
        coringa340: '-',
        coringa341: '-',
        coringa342: '-',
        _183: '',
        coringa343: '-',
        coringa344: '-',
        coringa345: '-',
        coringa346: '-',
        coringa347: '-',
        coringa348: '-',
        coringa349: '-',
        coringa350: '-',
        coringa351: '-',
        coringa352: '-',
        coringa353: '-',
        coringa354: '-',
        coringa355: '-',
        coringa356: '-',
        coringa357: '-',
        _184: '',
        coringa358: '-',
        coringa359: '-',
        coringa360: '-',
        coringa361: '-',
        coringa362: '-',
        coringa363: '-',
        coringa364: '-',
        coringa365: '-',
        coringa366: '-',
        coringa367: '-',
        coringa368: '-',
        coringa369: '-',
        coringa370: '-',
        coringa371: '-',
        coringa372: '-',
        _185: '',
        coringa373: '-',
        coringa374: '-',
        coringa375: '-',
        coringa376: '-',
        coringa377: '-',
        coringa378: '-',
        coringa379: '-',
        coringa380: '-',
        coringa381: '-',
        coringa382: '-',
        coringa383: '-',
        coringa384: '-',
        coringa385: '-',
        coringa386: '-',
        coringa387: '-',
        coringa388: '-',
        coringa389: '-',
        coringa390: '-',
        coringa391: '-',
        coringa392: '-',
        coringa393: '-',
        coringa394: '-',
        coringa395: '-',
        coringa396: '-',
        coringa397: '-',
        _186: '',
        coringa398: '-',
        coringa399: '-',
        coringa400: '-',
        coringa401: '-',
        coringa402: '-',
        coringa403: '-',
        coringa404: '-',
        coringa405: '-',
        coringa406: '-',
        coringa407: '-',
        coringa408: '-',
        coringa409: '-',
        coringa410: '-',
        coringa411: '-',
        coringa412: '-',
        coringa413: '-',
        coringa414: '-',
        coringa415: '-',
        coringa416: '-',
        coringa417: '-',
        coringa418: '-',
        coringa419: '-',
        coringa420: '-',
        coringa421: '-',
        coringa422: '-',
        _187: '',
        vivoPlayBasico: '-',
        vivoPlayEssencial: '-',
        vivoPlayAvancado: '-',
        vivoPlayCompleto: '-',
        pacoteStandardHdFibra: '-',
        pacoteSuperHdFibra: '-',
        pacoteUltimateHdFibra: '-',
        pacoteUltraHdFibra: '-',
        pacoteFullHdFibra: '-',
        coringa423: '-',
        coringa424: '-',
        coringa425: '-',
        coringa426: '-',
        coringa427: '-',
        coringa428: '-',
        coringa429: '-',
        coringa430: '-',
        coringa431: '-',
        coringa432: '-',
        coringa433: '-',
        coringa434: '-',
        coringa435: '-',
        coringa436: '-',
        coringa437: '-',
        coringa438: '-',
        _188: '',
        premiere: '-',
        telecine: '-',
        telecineLigthHd: '-',
        hbo: '-',
        esportes: '-',
        combate: '-',
        dogTv: '-',
        curta: '-',
        internacional: '-',
        superHot: '-',
        mantoman: '-',
        sexyHot: '-',
        coringa439: '-',
        coringa440: '-',
        coringa441: '-',
        coringa442: '-',
        coringa443: '-',
        coringa444: '-',
        coringa445: '-',
        coringa446: '-',
        coringa447: '-',
        coringa448: '-',
        coringa449: '-',
        coringa450: '-',
        coringa451: '-',
        _189: '',
        coringa452: '-',
        coringa453: '-',
        coringa454: '-',
        coringa455: '-',
        coringa456: '-',
        coringa457: '-',
        coringa458: '-',
        coringa459: '-',
        coringa460: '-',
        coringa461: '-',
        coringa462: '-',
        coringa463: '-',
        coringa464: '-',
        coringa465: '-',
        coringa466: '-',
        coringa467: '-',
        coringa468: '-',
        coringa469: '-',
        coringa470: '-',
        coringa471: '-',
        coringa472: '-',
        coringa473: '-',
        coringa474: '-',
        coringa475: '-',
        coringa476: '-',
        _190: '',
        coringa477: '-',
        coringa478: '-',
        coringa479: '-',
        coringa480: '-',
        coringa481: '-',
        coringa482: '-',
        coringa483: '-',
        coringa484: '-',
        coringa485: '-',
        coringa486: '-',
        coringa487: '-',
        coringa488: '-',
        coringa489: '-',
        coringa490: '-',
        coringa491: '-',
        coringa492: '-',
        coringa493: '-',
        coringa494: '-',
        coringa495: '-',
        coringa496: '-',
        coringa497: '-',
        coringa498: '-',
        coringa499: '-',
        coringa500: '-',
        coringa501: '-',
        _191: '',
        vivoPosIndividual7Gb1: '-',
        vivoFamiliaCompleta10Gb1: '-',
        vivoFamiliaCompleta20Gb1: '-',
        coringa502: '-',
        coringa503: '-',
        coringa504: '-',
        coringa505: '-',
        coringa506: '-',
        coringa507: '-',
        coringa508: '-',
        coringa509: '-',
        coringa510: '-',
        coringa511: '-',
        coringa512: '-',
        coringa513: '-',
        coringa514: '-',
        coringa515: '-',
        coringa516: '-',
        coringa517: '-',
        coringa518: '-',
        coringa519: '-',
        coringa520: '-',
        coringa521: '-',
        coringa522: '-',
        coringa523: '-',
        _192: '',
        coringa524: '-',
        coringa525: '-',
        coringa526: '-',
        coringa527: '-',
        coringa528: '-',
        coringa529: '-',
        coringa530: '-',
        coringa531: '-',
        coringa532: '-',
        coringa533: '-',
        coringa534: '-',
        coringa535: '-',
        coringa536: '-',
        coringa537: '-',
        coringa538: '-',
        coringa539: '-',
        coringa540: '-',
        coringa541: '-',
        coringa542: '-',
        coringa543: '-',
        coringa544: '-',
        coringa545: '-',
        coringa546: '-',
        coringa547: '-',
        coringa548: '-',
        _193: '',
        vivoPosIndividual7Gb2: '-',
        vivoFamiliaCompleta10Gb2: '-',
        vivoFamiliaCompleta20Gb2: '-',
        coringa549: '-',
        coringa550: '-',
        coringa551: '-',
        coringa552: '-',
        coringa553: '-',
        coringa554: '-',
        coringa555: '-',
        coringa556: '-',
        coringa557: '-',
        coringa558: '-',
        coringa559: '-',
        coringa560: '-',
        coringa561: '-',
        coringa562: '-',
        coringa563: '-',
        coringa564: '-',
        coringa565: '-',
        coringa566: '-',
        coringa567: '-',
        coringa568: '-',
        coringa569: '-',
        coringa570: '-',
        _194: '',
        _195: '',
        _196: '',
        coringa571: '-',
        coringa572: '-',
        coringa573: '-',
        coringa574: '-',
        coringa575: '-',
        coringa576: '-',
        coringa577: '-',
        coringa578: '-',
        coringa579: '-',
        coringa580: '-',
        coringa581: '-',
        coringa582: '-',
        _197: '',
        _198: '',
        solo: '-',
        duo: '-',
        trio: '-',
    })

    pProcess.setAttribute('style', 'display: none')
    loader.setAttribute('style', 'display: none')

    exportToCSV(dataExport, 'dataFixa.csv');
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