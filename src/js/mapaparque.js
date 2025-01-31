const loader = document.getElementById("loader");
const pProcess = document.getElementById("pProcessando");
const name = document.getElementById("fileName");
const fileName = document.getElementById("labelFileInput");
const processarBtn = document.getElementById("processarBtn");
const fileInput = document.getElementById("fileInput");
let mapaParqueData = [];
let dispData = [];
let cnaeData = [];
let cepData = [];

fetch("../public/cep.json").then((response) => {
  response.json().then((e) => {
    console.log(e);

    e.forEach((el) => {
      faixaCep.push(el.col3);
      faixaCep.push(el.col4);
    });

    faixaCep.sort((a, b) => {
      return parseInt(a) - parseInt(b);
    });

    cepData = e;
  });
});

fetch("../public/cnae.json").then((response) => {
  response.json().then((e) => {
    cnaeData = e;
  });
});

let fixaData = [];
let fixaFTTH = [];
let fixaFTTC = [];
let faixaCep = [];
let tiData = [];
let movelData = [];
let avancadaData = [];
let suspensoData = [];
let recomendacaoMovelData = [];
let recomendacaoFixaData = [];
let sipData = [];
let vozData = [];
let M2MData = [];
let newCSVContent = "";

let cnaeFile;
let cepFile;
let recMovelFile;
let vozFile;
let sipFile;
let suspensoFile;
let tiFile;
let avancadaFile;
let movelFile;
let m2mFile;
let fixaFile;
let mapaDispFile;
let mapaParqueFile;
let filesArray = [];

function contarOcorrenciasEPrimeiroRegistro(array) {
  return array.reduce((resultado, item) => {
    const cnpj = item.CNPJ_CLIENTE;

    // Se o CNPJ ainda não foi registrado, cria o registro inicial
    if (!resultado[cnpj]) {
      resultado[cnpj] = {
        quantidade: 1,
        primeiroRegistro: item.CODCLI,
      };
    } else {
      // Se já existe, apenas incrementa a quantidade
      resultado[cnpj].quantidade += 1;
    }

    return resultado;
  }, {});
}

fileInput.addEventListener(
  "change",
  (e) => {
    console.log("Disparou o evento de change"); // Verifique se aparece 2x
    handleFiles(e.target.files);
  },
  { once: true }
); // Adiciona o listener apenas uma vez

// Função para ler arquivos Excel usando SheetJS
async function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Lendo apenas a primeira planilha
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      resolve({ fileName: file.name, data: jsonData });
    };

    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

// Manipular múltiplos arquivos
async function handleFiles(fileList) {
  console.log("Iniciando processamento..."); // Verifique se aparece 2x

  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const files = Array.from(fileList);

  files.sort((a, b) => {
    const order = [
      "recMovel",
      "voz",
      "sip",
      "suspenso",
      "ti",
      "avancada",
      "movel",
      "m2m",
      "fixa",
      "mapadisp",
      "mapaparque",
    ]; // Ordem desejada
    const aIndex = order.findIndex((keyword) => a.name.includes(keyword));
    const bIndex = order.findIndex((keyword) => b.name.includes(keyword));

    return (
      (aIndex === -1 ? order.length : aIndex) -
      (bIndex === -1 ? order.length : bIndex)
    );
  });

  for (const file of files) {
    console.log(`Lendo arquivo: ${file.name}`); // Verifique duplicidade

    const { fileName, data } = await readExcelFile(file);

    if (fileName.includes("recMovel")) {
      recMovel(data);
    } else if (fileName.includes("voz")) {
      voz(data);
    } else if (fileName.includes("sip")) {
      sip(data);
    } else if (fileName.includes("suspenso")) {
      suspenso(data);
    } else if (fileName.includes("ti")) {
      ti(data);
    } else if (fileName.includes("avancada")) {
      avancada(data);
    } else if (fileName.includes("movel")) {
      movel(data);
    } else if (fileName.includes("m2m")) {
      m2m(data);
    } else if (fileName.includes("fixa")) {
      fixa(data);
    } else if (fileName.includes("mapadisp")) {
      mapaDisp(data);
    } else if (fileName.includes("mapaparque")) {
      mapaParque(data);
    } else {
      console.warn(`Arquivo ${fileName} não reconhecido.`);
    }
  }
}

// ===========  Arquivo Mapa Parque  ===========
async function mapaParque(arquivo) {
  console.log(arquivo);
  arquivo.forEach((e) => {
    // Zerar propriedades
    e.DS_CIDADE = 0;
    e.DS_DISPONIBILIDADE = 0;

    // Criar propriedades
    let tipoClienteBasica = "";
    let flgDispGpon = 0;
    let agrupamentoCnaeMacroAtual = "";
    let agrupamentoCnaeMicroAtual = "";
    let uf = 0;
    let faturamento = 0;

    // Verificação
    if (e.TP_PRODUTO !== 0) {
      if (e.TP_PRODUTO.includes("AVANÇADO")) {
        e.TP_PRODUTO = e.TP_PRODUTO.replace("AVANÇADO", "AVANÇADA");
      }
    }

    let valorCep = e.NR_CEP.toString();
    if (valorCep !== undefined) {
      e.NR_CEP = valorCep.replace("-", "");
    }

    cnaeData.forEach((eC) => {
      if (e.DS_CNAE === eC.COD_CNAE) {
        agrupamentoCnaeMacroAtual = eC.AGRUPAMENTOCNAEMACROATUAL;
        agrupamentoCnaeMicroAtual = eC.AGRUPAMENTOCNAEMICROATUAL;
      }
    });

    const indexFTTH = fixaFTTH.findIndex((eB) => eB.DOCUMENTO === e.NR_CNPJ);
    const indexFTTC = fixaFTTC.findIndex((eB) => eB.DOCUMENTO === e.NR_CNPJ);

    if (e.TP_PRODUTO !== 0) {
      if (indexFTTH !== -1) {
        tipoClienteBasica = "FTTH";
      } else if (indexFTTC !== -1) {
        tipoClienteBasica = "FTTC";
      } else if (e.TP_PRODUTO.includes("BASICA")) {
        tipoClienteBasica = "1P_VOZ_SOLO";
      } else {
        tipoClienteBasica = "SEM PRODUTO";
      }
    } else {
      tipoClienteBasica = "SEM PRODUTO";
    }

    const result = dispData.findIndex((eB) => eB.NR_CNPJ === e.NR_CNPJ);
    if (result !== -1) {
      e.DS_DISPONIBILIDADE = dispData[result].DS_DISPONIBILIDADE;
      flgDispGpon = 1;
    }

    let valor = e.NR_CEP;

    let maisProximo = faixaCep.reduce(function (anterior, corrente) {
      return Math.abs(corrente - valor) < Math.abs(anterior - valor)
        ? corrente
        : anterior;
    });

    cepData.forEach((eB) => {
      if (maisProximo === eB.col3 || maisProximo === eB.col4) {
        uf = eB.col1;
        e.DS_CIDADE = eB.col2;
      }
    });

    fixaData.forEach((eFixa) => {
      if (e.NR_CNPJ === eFixa.DOCUMENTO) {
        faturamento += parseFloat(eFixa.VL_FAT_BRUTO);
      }
    });
    avancadaData.forEach((eAvancada) => {
      if (e.NR_CNPJ === eAvancada.NRCNPJ) {
        faturamento += parseFloat(eAvancada.VALOR_TOTAL);
      }
    });
    movelData.forEach((eMovel) => {
      if (e.NR_CNPJ === eMovel.CNPJ_CLIENTE) {
        faturamento += parseFloat(eMovel.FAT_MEDIO_03_MESES);
      }
    });
    tiData.forEach((eTi) => {
      if (e.NR_CNPJ === eTi.CNPJ) {
        faturamento += parseFloat(eTi.VALOR);
      }
    });

    mapaParqueData.push({
      COD_CLIENTE: e.COD_CLIENTE,
      NR_CNPJ: e.NR_CNPJ,
      QT_MOVEL_TERM: e.QT_MOVEL_TERM,
      QT_MOVEL_PEN: e.QT_MOVEL_PEN,
      QT_MOVEL_M2M: e.QT_MOVEL_M2M,
      QT_MOVEL_FWT: e.QT_MOVEL_FWT,
      QT_BASICA_FWT: e.QT_BASICA_FWT,
      QT_BASICA_TERM_FIBRA: e.QT_BASICA_TERM_FIBRA,
      QT_BASICA_TERM_METALICO: e.QT_BASICA_TERM_METALICO,
      QT_BASICA_BL: e.QT_BASICA_BL,
      QT_BASICA_TV: e.QT_BASICA_TV,
      QT_BASICA_OUTROS: e.QT_BASICA_OUTROS,
      QT_BASICA_LINAS: e.QT_BASICA_LINAS,
      QT_AVANCADA_DDR: e.QT_AVANCADA_DDR,
      QT_AVANCADA_RI: e.QT_AVANCADA_RI,
      QT_AVANCADA_RDSI: e.QT_AVANCADA_RDSI,
      QT_AVANCADA_TERM: e.QT_AVANCADA_TERM,
      QT_AVANCADA_VOX: e.QT_AVANCADA_VOX,
      QT_AVANCADA_SIP: e.QT_AVANCADA_SIP,
      QT_AVANCADA_DADOS: e.QT_AVANCADA_DADOS,
      QT_VIVO_TECH: e.QT_VIVO_TECH,
      QT_OFFICE_365: e.QT_OFFICE_365,
      VL_CAR_MOVEL: 0,
      QT_CAR_MOVEL: 0,
      VL_CAR_FIXA: 0,
      QT_CAR_FIXA: 0,
      DS_DISPONIBILIDADE: e.DS_DISPONIBILIDADE,
      FLG_MEI: e.FLG_MEI,
      FLG_NAO_PERTURBE: e.FLG_NAO_PERTURBE,
      SEMAFORO_SERASA: e.SEMAFORO_SERASA,
      MENSAGEM_RETORNO_SERASA: e.MENSAGEM_RETORNO_SERASA,
      PROPENSAO_CONCORRENCIA_PMG: e.PROPENSAO_MOVEL_AVANCADA,
      DS_CNAE: e.DS_CNAE,
      DS_ATIVIDADE_ECONOMICA: e.DS_ATIVIDADE_ECONOMICA,
      ADABAS_MOVEL: e.ADABASMOVEL,
      ADABAS_FIXA: e.ADABASFIXA,
      EXCLUIR_PARQUE: "NAO",
      TORRE_PRODUTOS: e.TP_PRODUTO,
      PERFIL_PARQUE_MOVEL: "%PERFIL_PARQUE_MOVEL",
      PERFIL_PARQUE_BASICA: "%PERFIL_PARQUE_BASICA",
      PERFIL_PARQUE_AVANCADA: "%PERFIL_PARQUE_AVANCADA",
      BOAS_VINDAS: 0,
      TIPO_CLIENTE_BASICA: tipoClienteBasica,
      FLG_CID_DISP_VVN: e.FLG_CID_DISP_VVN,
      FLG_ERB: e.FLG_ERB,
      FLG_DISP_GPON: flgDispGpon,
      AGRUPAMENTO_CNAE_MACRO_ATUAL: agrupamentoCnaeMacroAtual,
      AGRUPAMENTO_CNAE_MICRO_ATUAL: agrupamentoCnaeMicroAtual,
      CLUSTER: e.CLUSTER,
      OFERTA_1: e.PRIMEIRA_OFERTA,
      OFERTA_2: e.SEGUNDA_OFERTA,
      OFERTA_3: e.TERCEIRA_OFERTA,
      OFERTA_4: 0,
      OFERTA_5: 0,
      NM_CONTATO_SFA: e.NM_CONTATO_SFA,
      EMAIL_CONTATO_PRINCIPAL_SFA: e.EMAIL_CONTATO_PRINCIPAL_SFA,
      CELULAR_CONTATO_PRINCIPAL_SFA: e.CELULAR_CONTATO_PRINCIPAL_SFA,
      PAPEL_CONTATO_SFA: e.PAPEL_CONTATO_SFA,
      FLG_DOMINIO_PUBLICO_SFA: e.FLG_DOMINIO_PUBLICO_SFA,
      TLFN_1: e.TLFN_1,
      TLFN_2: e.TLFN_2,
      TLFN_3: e.TLFN_3,
      TLFN_4: e.TLFN_4,
      TLFN_5: e.TLFN_5,
      TLFN_6: e.TLFN_6,
      TLFN_7: e.TLFN_7,
      TLFN_8: e.TLFN_8,
      TEL_COMERCIAL_SIEBEL: e.TEL_COMERCIAL_SIEBEL,
      TEL_CELULAR_SIEBEL: e.TEL_CELULAR_SIEBEL,
      TEL_RESIDENCIAL_SIEBEL: e.TEL_RESIDENCIAL_SIEBEL,
      UF: uf,
      MUNICIPIO: e.DS_CIDADE,
      FATURAMENTO: faturamento,
    });
  });
  verificar();
}

function objectToCsvRow(obj) {
  const values = Object.values(obj);
  return values.map((value) => `"${value}"`).join(",") + "\n";
}

// ===========  Arquivo Suspenso  ===========
async function suspenso(arquivo) {
  console.log(arquivo);
  suspensoData = arquivo;
  console.log(suspensoData);
}

// ===========  Arquivo M2M  ===========
async function m2m(arquivo) {
  console.log(arquivo);
  arquivo.forEach((e) => {
    let dataPrimeiraAtivacao = numeroInteiroParaData(e.DT_PRIMEIRA_ATIVACAO);

    let indexSuspensao = suspensoData.findIndex(
      (element) => element.CNPJ_CLIENTE === e.CNPJ_CLIENTE
    );
    let linhaSuspensa = 0;
    let linhaSuspensaStatus = 0;

    if (indexSuspensao !== -1) {
      linhaSuspensaStatus = suspensoData[indexSuspensao].STATUS;
      linhaSuspensa = "Sim";
    } else {
      linhaSuspensaStatus = 0;
      linhaSuspensa = "Nao";
    }

    movelData.push({
      CODCLI: e.CODCLI,
      CNPJ_CLIENTE: e.CNPJ_CLIENTE,
      NR_TELEFONE: e.NR_TELEFONE,
      UF: e.UF,
      PLANO: e.PLANO,
      M: 0,
      FIDELIZADO: e.FIDELIZADO,
      DT_PRIMEIRA_ATIVACAO: `${dataPrimeiraAtivacao.getDate() + 1}/${
        dataPrimeiraAtivacao.getMonth() + 1
      }/${dataPrimeiraAtivacao.getFullYear()}`,
      PLANO_CONTA: e.PLANO_CONTA,
      PLANO_LINHA: e.PLANO_LINHA,
      LINHA_SUSPENSA: linhaSuspensa,
      LINHA_SUSPENSA_STATUS: linhaSuspensaStatus,
      FAT_MEDIO_03_MESES: 0,
      QT_DIAS_TRAF_DADOS: 0,
      QT_MB_TRAF_DADOS: 0,
    });
  });
}

// ===========  Arquivo Recomendação Móvel  ===========
async function recMovel(arquivo) {
  console.log(arquivo);
  recomendacaoMovelData = arquivo;
  console.log(recomendacaoMovelData);
}

// ===========  Arquivo Sip  ===========
async function sip(arquivo) {
  console.log(arquivo);
  sipData = arquivo;
  console.log(sipData);
  sipData.forEach((e) => {
    let valor = e.END_CEP;

    let maisProximo = faixaCep.reduce(function (anterior, corrente) {
      return Math.abs(corrente - valor) < Math.abs(anterior - valor)
        ? corrente
        : anterior;
    });

    cepData.forEach((eB) => {
      if (maisProximo === eB.col3 || maisProximo === eB.col4) {
        e.END_UF = eB.col1;
        e.END_CIDADE = eB.col2;
      }
    });

    let data_criacao = 0;
    let dataFinal = 0;
    if (typeof e.DATA_CRIACAO === "string") {
      data_criacao = new Date(e.DATA_CRIACAO.slice(0, -13));
      dataFinal = new Date(data_criacao);
    } else if (typeof e.DATA_CRIACAO === "number") {
      data_criacao = numeroInteiroParaData(
        String(e.DATA_CRIACAO).split(".")[0]
      );
      dataFinal = new Date(data_criacao);
    }

    dataFinal.setFullYear(data_criacao.getFullYear() + 3);
    avancadaData.push({
      CD_PESSOA: e.COD_CLIENTE,
      NMLINHANEGOCIO: "SIP",
      NMLINHAPRODUTO: "PLANO FLEX",
      VELOCIDADE: e.QTD_CANAIS,
      NRCNPJ: e.CNPJ,
      NUM_MESES_PRAZO_CONTRATUAL: 36,
      DT_ASS: data_criacao,
      DT_FIM_PRAZO_CONTRATUAL: dataFinal,
      VALOR_TOTAL: e.VL_TOTAL,
      END_CIDADE: e.END_CIDADE,
      END_ESTADO: e.END_UF,
      END_CEP: e.END_CEP,
      PERCENTUAL_DE_CONTRATO: 100,
      TERMINAL_TRONCO: e.TERMINAL,
    });
  });
}

// ===========  Arquivo Voz  ===========
async function voz(arquivo) {
  console.log(arquivo);
  vozData = arquivo;
  console.log(vozData);

  vozData.forEach((e) => {
    const string = e.TP_PRODUTO;
    let match = "";
    if (string.includes(" ")) {
      [, match] = string.match(/(\S+) /) || [];
    } else {
      match = e.TP_PRODUTO;
    }

    let valor = e.NR_CEP;

    let maisProximo = faixaCep.reduce(function (anterior, corrente) {
      return Math.abs(corrente - valor) < Math.abs(anterior - valor)
        ? corrente
        : anterior;
    });

    cepData.forEach((eB) => {
      if (maisProximo === eB.col3 || maisProximo === eB.col4) {
        e.UF = eB.col1;
        e.DS_MUNICIPIO = eB.col2;
      }
    });

    avancadaData.push({
      CD_PESSOA: e.CD_PESSOA,
      NMLINHANEGOCIO: match,
      NMLINHAPRODUTO: e.TP_PRODUTO,
      VELOCIDADE: e.QTD_ACESSO,
      NRCNPJ: 0,
      NUM_MESES_PRAZO_CONTRATUAL: 0,
      DT_ASS: 0,
      DT_FIM_PRAZO_CONTRATUAL: 0,
      VALOR_TOTAL: e.VL_TOTAL,
      END_CIDADE: e.DS_MUNICIPIO,
      END_ESTADO: e.UF,
      END_CEP: e.NR_CEP,
      PERCENTUAL_DE_CONTRATO: 100,
      TERMINAL_TRONCO: e.TRONCO,
    });
  });
}

// ===========  Arquivo Movel  ===========
async function movel(arquivo) {
  console.log(arquivo);
  arquivo.forEach((e) => {
    let linhaSuspensa = "Nao";
    let linhaSuspensaStatus = 0;

    // Renomear
    delete Object.assign(e, { ["M"]: e["M_RECOMENDACAO"] })["M_RECOMENDACAO"];

    if (parseInt(e.M) >= 17) {
      e.FIDELIZADO = "Nao Fidelizado";
    } else if (parseInt(e.M < 17)) {
      e.FIDELIZADO = "Fidelizado";
    }

    let indexSusp = suspensoData.findIndex(
      (element) => element.CNPJ_CLIENTE === e.CNPJ_CLIENTE
    );
    if (indexSusp !== -1) {
      linhaSuspensaStatus = suspensoData[indexSusp].STATUS;
      linhaSuspensa = "Sim";
    }

    let tdata = 0;
    if (e.DT_PRIMEIRA_ATIVACAO !== 0) {
      let dataPrimeiraAtivacao = numeroInteiroParaData(e.DT_PRIMEITA_ATIVACAO);
      tdata = `${dataPrimeiraAtivacao.getDate() + 1}/${
        dataPrimeiraAtivacao.getMonth() + 1
      }/${dataPrimeiraAtivacao.getFullYear()}`;
    }

    let indexRec = recomendacaoMovelData.findIndex(
      (element) => element.NR_DOCUMENTO === e.CNPJ_CLIENTE
    );
    if (indexRec !== -1) {
      e.PLANO_CONTA = recomendacaoMovelData[indexRec].PLANO_RECOMENDADO;
      e.PLANO_LINHA = recomendacaoMovelData[indexRec].PLANO_RECOMENDADO_UP;
    } else {
      e.PLANO_CONTA = 0;
      e.PLANO_LINHA = 0;
    }

    movelData.push({
      CODCLI: e.CODCLI,
      CNPJ_CLIENTE: e.CNPJ_CLIENTE,
      NR_TELEFONE: e.NR_TELEFONE,
      UF: e.UF,
      PLANO: e.PLANO,
      M: e.M,
      FIDELIZADO: e.FIDELIZADO,
      DT_PRIMEIRA_ATIVACAO: tdata,
      PLANO_CONTA: e.PLANO_CONTA,
      PLANO_LINHA: e.PLANO_LINHA,
      LINHA_SUSPENSA: linhaSuspensa,
      LINHA_SUSPENSA_STATUS: linhaSuspensaStatus,
      FAT_MEDIO_03_MESES: e.FAT_MEDIO_03_MESES,
      QT_DIAS_TRAF_DADOS: e.QT_DIAS_TRAF_DADOS,
      QT_MB_TRAF_DADOS: e.QT_MB_TRAF_DADOS,
    });
  });
}

// ===========  Arquivo Avançada  ===========
async function avancada(arquivo) {
  console.log(arquivo);
  arquivo.forEach((e) => {
    // Renomear Propriedade
    delete Object.assign(e, { ["TERMINAL_TRONCO"]: e["SISTEMA_ORIGEM"] })[
      "SISTEMA_ORIGEM"
    ];

    // Zerar
    let endCidade = 0;
    let endEstado = 0;

    let tInicial = numeroInteiroParaData(e.DT_ASS);
    let tFinal = numeroInteiroParaData(e.DT_FIM_PRAZO_CONTRATUAL);
    let hoje = new Date();

    let dataInicial = `${tInicial.getDate()}/${
      tInicial.getMonth() + 1
    }/${tInicial.getFullYear()}`;
    let dataFinal = `${tFinal.getDate()}/${
      tFinal.getMonth() + 1
    }/${tFinal.getFullYear()}`;

    let result1 = tFinal.getTime() - tInicial.getTime();
    let result2 = hoje.getTime() - tInicial.getTime();
    let result3 = (result2 / result1) * 100;

    let valor = e.END_CEP;

    let maisProximo = faixaCep.reduce(function (anterior, corrente) {
      return Math.abs(corrente - valor) < Math.abs(anterior - valor)
        ? corrente
        : anterior;
    });

    cepData.forEach((eB) => {
      if (maisProximo === eB.col3 || maisProximo === eB.col4) {
        endEstado = eB.col1;
        endCidade = eB.col2;
      }
    });

    avancadaData.push({
      CD_PESSOA: e.CD_PESSOA,
      NMLINHANEGOCIO: "SIP",
      NMLINHAPRODUTO: "PLANO FLEX",
      VELOCIDADE: e.VELOCIDADE,
      NRCNPJ: e.NRCNPJ,
      NUM_MESES_PRAZO_CONTRATUAL: e.NUM_MESES_PRAZO_CONTRATUAL,
      DT_ASS: dataInicial,
      DT_FIM_PRAZO_CONTRATUAL: dataFinal,
      VALOR_TOTAL: e.VALOR_TOTAL,
      END_CIDADE: endCidade,
      END_ESTADO: endEstado,
      END_CEP: e.END_CEP,
      PERCENTUAL_DE_CONTRATO: parseInt(result3),
      TERMINAL_TRONCO: e.TERMINAL_TRONCO,
    });
  });
}

// Função para converter número inteiro em data
function numeroInteiroParaData(numero) {
  var dataBase = new Date(1900, 0, -1); // Data base do Excel (1º de janeiro de 1900)
  var data = new Date(dataBase.getTime() + numero * 24 * 60 * 60 * 1000);
  return data;
}

// ===========  Arquivo Ti  ===========
async function ti(arquivo) {
  console.log(arquivo);
  arquivo.forEach((e) => {
    let tInicio = numeroInteiroParaData(e.INICIO);
    let tFim = numeroInteiroParaData(e.FIM);

    e.INICIO = `${tInicio.getDate() + 1}/${
      tInicio.getMonth() + 1
    }/${tInicio.getFullYear()}`;
    e.FIM = `${tFim.getDate() + 1}/${
      tFim.getMonth() + 1
    }/${tFim.getFullYear()}`;

    e.PERCENTUAL_CONTRATO = `${parseFloat(e.PERCENTUAL_CONTRATO * 100).toFixed(
      2
    )}%`;
  });
  tiData = arquivo;
  console.log(tiData);
}

// ===========  Arquivo Fixa  ===========
async function fixa(arquivo) {
  console.log(arquivo);
  fixaData = arquivo;
  fixaData.forEach((e) => {
    if (e.DS_PRODUTO_TECNOLOGIA === "BANDA LARGA - FIBRA") {
      fixaFTTH.push(e);
    } else if (e.DS_PRODUTO_TECNOLOGIA === "BANDA LARGA - METALICO") {
      fixaFTTC.push(e);
    }

    if (e.DT_INI_FDLZ != 0) {
      let tIniFdlz = numeroInteiroParaData(e.DT_INI_FDLZ);
      let tFimFdlz = numeroInteiroParaData(e.DT_FIM_FDLZ);

      e.DT_INI_FDLZ = `${tIniFdlz.getDate() + 1}/${
        tIniFdlz.getMonth() + 1
      }/${tIniFdlz.getFullYear()}`;
      e.DT_FIM_FDLZ = `${tFimFdlz.getDate() + 1}/${
        tFimFdlz.getMonth() + 1
      }/${tFimFdlz.getFullYear()}`;
    } else {
      e.DT_INI_FDLZ = 0;
      e.DT_FIM_FDLZ = 0;
    }
  });
  console.log(fixaData);
}

// ===========  Arquivo Disp  ===========
async function mapaDisp(arquivo) {
  console.log(arquivo);
  dispData = arquivo;
  console.log(dispData);
}

async function verificar() {
  avancadaData.forEach((e) => {
    if (e.NRCNPJ === 0) {
      let index = mapaParqueData.findIndex(
        (element) => element.COD_CLIENTE === e.CD_PESSOA
      );
      if (index !== -1) {
        e.NRCNPJ = mapaParqueData[index].NR_CNPJ;
      } else {
        e.NRCNPJ = 0;
      }
    }
  });
  console.log("Verificou: CNPJ Faltante");

  let arrayQuantidade = contarOcorrenciasEPrimeiroRegistro(movelData);
  Object.entries(arrayQuantidade).forEach(([key, value]) => {
    let indexMapa = mapaParqueData.findIndex(
      (element) => element.NR_CNPJ === key
    );
    if (indexMapa === -1) {
      mapaParqueData.push({
        COD_CLIENTE: value.CODCLI,
        NR_CNPJ: value,
        QT_MOVEL_TERM: value.quantidade,
        QT_MOVEL_PEN: 0,
        QT_MOVEL_M2M: 0,
        QT_MOVEL_FWT: 0,
        QT_BASICA_FWT: 0,
        QT_BASICA_TERM_FIBRA: 0,
        QT_BASICA_TERM_METALICO: 0,
        QT_BASICA_BL: 0,
        QT_BASICA_TV: 0,
        QT_BASICA_OUTROS: 0,
        QT_BASICA_LINAS: 0,
        QT_AVANCADA_DDR: 0,
        QT_AVANCADA_RI: 0,
        QT_AVANCADA_RDSI: 0,
        QT_AVANCADA_TERM: 0,
        QT_AVANCADA_VOX: 0,
        QT_AVANCADA_SIP: 0,
        QT_AVANCADA_DADOS: 0,
        QT_VIVO_TECH: 0,
        QT_OFFICE_365: 0,
        VL_CAR_MOVEL: 0,
        QT_CAR_MOVEL: 0,
        VL_CAR_FIXA: 0,
        QT_CAR_FIXA: 0,
        DS_DISPONIBILIDADE: 0,
        FLG_MEI: 0,
        FLG_NAO_PERTURBE: 0,
        SEMAFORO_SERASA: 0,
        MENSAGEM_RETORNO_SERASA: 0,
        PROPENSAO_CONCORRENCIA_PMG: 0,
        DS_CNAE: 0,
        DS_ATIVIDADE_ECONOMICA: 0,
        ADABAS_MOVEL: 0,
        ADABAS_FIXA: 0,
        EXCLUIR_PARQUE: "NAO",
        TORRE_PRODUTOS: 0,
        PERFIL_PARQUE_MOVEL: 0,
        PERFIL_PARQUE_BASICA: 0,
        PERFIL_PARQUE_AVANCADA: 0,
        BOAS_VINDAS: 0,
        TIPO_CLIENTE_BASICA: 0,
        FLG_CID_DISP_VVN: 0,
        FLG_ERB: 0,
        FLG_DISP_GPON: 0,
        AGRUPAMENTO_CNAE_MACRO_ATUAL: 0,
        AGRUPAMENTO_CNAE_MICRO_ATUAL: 0,
        CLUSTER: 0,
        OFERTA_1: 0,
        OFERTA_2: 0,
        OFERTA_3: 0,
        OFERTA_4: 0,
        OFERTA_5: 0,
        NM_CONTATO_SFA: 0,
        EMAIL_CONTATO_PRINCIPAL_SFA: 0,
        CELULAR_CONTATO_PRINCIPAL_SFA: 0,
        PAPEL_CONTATO_SFA: 0,
        FLG_DOMINIO_PUBLICO_SFA: 0,
        TLFN_1: 0,
        TLFN_2: 0,
        TLFN_3: 0,
        TLFN_4: 0,
        TLFN_5: 0,
        TLFN_6: 0,
        TLFN_7: 0,
        TLFN_8: 0,
        TEL_COMERCIAL_SIEBEL: 0,
        TEL_CELULAR_SIEBEL: 0,
        TEL_RESIDENCIAL_SIEBEL: 0,
        UF: 0,
        MUNICIPIO: 0,
        FATURAMENTO: 0,
      });
    }
  });

  mapaParqueData.forEach((e) => {
    let eAvancada = avancadaData.find(
      (element) => element.NRCNPJ === e.NR_CNPJ
    );
    if (eAvancada !== undefined) {
      let percent = eAvancada.PERCENTUAL_DE_CONTRATO;
      if (percent < 50) {
        e.PERFIL_PARQUE_AVANCADA = "CONTRATO < 50%";
      } else if (percent < 75) {
        e.PERFIL_PARQUE_AVANCADA = "CONTRATO > 50% E < 75%";
      } else {
        e.PERFIL_PARQUE_AVANCADA = "CONTRATO >= 75%";
      }
    } else {
      e.PERFIL_PARQUE_AVANCADA = "SEM PRODUTOS AVANÇADA";
    }

    function compare(a, b) {
      if (a.M < b.M) return -1;
      if (a.M > b.M) return 1;
      return 0;
    }

    let eBasica = fixaData.filter((element) => element.DOCUMENTO === e.NR_CNPJ);
    if (eBasica.length !== 0) {
      eBasica.sort(compare);
      let menor = eBasica[0];
      let maior = eBasica[eBasica.length - 1];

      if (maior.M < 7 && menor.M >= 0) {
        e.PERFIL_PARQUE_BASICA = "PRODUTOS M < 7";
      } else if (maior.M < 17 && menor.M > 6) {
        e.PERFIL_PARQUE_BASICA = "PRODUTOS M 7 A 16";
      } else if (maior.M > 24 && menor.M > 24) {
        e.PERFIL_PARQUE_BASICA = "PRODUTOS M >= 25";
      } else {
        e.PERFIL_PARQUE_BASICA = "PRODUTOS MIX";
      }
    } else {
      e.PERFIL_PARQUE_BASICA = "SEM PRODUTOS BASICA";
    }

    let eMovel = movelData.filter(
      (element) => element.CNPJ_CLIENTE === e.NR_CNPJ
    );
    if (eMovel.length !== 0) {
      eMovel.sort(compare);
      let menor = eMovel[0];
      let maior = eMovel[eMovel.length - 1];

      if (maior.M < 7 && menor.M >= 0) {
        e.PERFIL_PARQUE_MOVEL = "PRODUTOS M < 7";
      } else if (maior.M < 17 && menor.M > 6) {
        e.PERFIL_PARQUE_MOVEL = "PRODUTOS M 7 A 16";
      } else if (maior.M < 25 && menor.M > 16) {
        e.PERFIL_PARQUE_MOVEL = "PRODUTOS M 17 A 24";
      } else if (maior.M > 24 && menor.M > 24) {
        e.PERFIL_PARQUE_MOVEL = "PRODUTOS M >= 25";
      } else {
        e.PERFIL_PARQUE_MOVEL = "PRODUTOS MIX";
      }
    } else {
      e.PERFIL_PARQUE_MOVEL = "SEM LINHA TERMINAL";
    }
  });
  console.log("Verificou: Perfil");
  console.log("Terminou a verificação! Exporte o arquivo");

  exportToExcel();

  loader.setAttribute("style", "display: none");
  pProcess.setAttribute("style", "display: none");
}

// ===========  Criar e Exportar Novo Arquivo  ===========
function exportToExcel() {
  // Cria uma nova worksheet a partir dos dados filtrados
  const wsMapaParque = XLSX.utils.json_to_sheet(mapaParqueData);
  const wsTi = XLSX.utils.json_to_sheet(tiData);
  const wsFixa = XLSX.utils.json_to_sheet(fixaData);
  const wsMovel = XLSX.utils.json_to_sheet(movelData);
  const wsAvancada = XLSX.utils.json_to_sheet(avancadaData);
  const wsRecomendacaoFixa = XLSX.utils.json_to_sheet(recomendacaoFixaData);

  // Cria um novo workbook e adiciona a worksheet a ele
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, wsMapaParque, "Mapa Parque");
  XLSX.utils.book_append_sheet(workbook, wsMovel, "Parque Móvel");
  XLSX.utils.book_append_sheet(workbook, wsFixa, "Parque Básica");
  XLSX.utils.book_append_sheet(workbook, wsAvancada, "Parque de Avançada");
  XLSX.utils.book_append_sheet(workbook, wsTi, "Parque TI");
  XLSX.utils.book_append_sheet(workbook, wsRecomendacaoFixa, "Recomendações");

  // Exporta o workbook como um arquivo XLSX
  XLSX.writeFile(workbook, "mapaParque.xlsx", { bookType: "xlsx" });
}
