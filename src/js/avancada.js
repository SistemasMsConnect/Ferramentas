const inputDados = document.getElementById("fileInputDados");
const inputSip = document.getElementById("fileInputSip");
const inputVoz = document.getElementById("fileInputVoz");

const labelDados = document.getElementById("labelFileInputDados");
const labelSip = document.getElementById("labelFileInputSip");
const labelVoz = document.getElementById("labelFileInputVoz");

let cepData = [];
let faixaCep = [];

let combinedDadosData = [];
let combinedSipData = [];
let combinedVozData = [];

let filesDadosProcessed = 0;
let filesSipProcessed = 0;
let filesVozProcessed = 0;

const dataExport = [];

const ufs = [
  "AC",
  "AL",
  "AP",
  "AM",
  "BA",
  "CE",
  "DF",
  "ES",
  "GO",
  "MA",
  "MT",
  "MS",
  "MG",
  "PA",
  "PB",
  "PR",
  "PE",
  "PI",
  "RJ",
  "RN",
  "RS",
  "RO",
  "RR",
  "SC",
  "SP",
  "SE",
  "TO",
];

// === === === Modificacao label === === ===
inputDados.addEventListener("change", (event) => {
  labelDados.textContent = event.target.files[0].name;
});
inputSip.addEventListener("change", (event) => {
  labelSip.textContent = event.target.files[0].name;
});
inputVoz.addEventListener("change", (event) => {
  labelVoz.textContent = event.target.files[0].name;
});
// === === === Modificacao label === === ===

// === === === Funcoes auxiliares === === ===
function numberToDate(data) {
  let dia = data.getDate();
  if (String(dia).length == 1) {
    dia = `0${dia}`;
  }
  let mes = data.getMonth() + 1;
  if (String(mes).length == 1) {
    mes = `0${mes}`;
  }
  let ano = data.getFullYear();

  let dataFormatada = `${ano}-${mes}-${dia}`;
  return dataFormatada;
}

function removerAcentos(texto) {
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function numeroInteiroParaData(numero) {
  var dataBase = new Date(1900, 0, -1); // Data base do Excel (1º de janeiro de 1900)
  var data = new Date(dataBase.getTime() + (numero + 1) * 24 * 60 * 60 * 1000);
  return data;
}

function exportToExcel(data, sheetname, filename) {
  // Cria uma nova worksheet a partir dos dados filtrados
  const worksheet1 = XLSX.utils.json_to_sheet(data);

  // Cria um novo workbook e adiciona a worksheet a ele
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet1, `${sheetname}`);

  // Exporta o workbook como um arquivo XLSX
  XLSX.writeFile(workbook, filename, { bookType: "xlsx" });
}
// === === === Funcoes auxiliares === === ===

// === === === Leitura dos arquivos === === ===
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

document.getElementById("processBtn").addEventListener("click", function () {
  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const filesDados = inputDados.files;
  const filesSip = inputSip.files;
  const filesVoz = inputVoz.files;

  if (
    filesDados.length === 0 ||
    filesSip.length === 0 ||
    filesVoz.length === 0
  ) {
    alert("Por favor, selecione pelo menos um arquivo para cada input.");
    return;
  }

  Array.from(filesDados).forEach((file) => {
    const reader = new FileReader();

    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Assume que o CSV tem apenas uma folha (sheet)
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Converte o sheet para JSON
      const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 0 });

      // csvData.forEach(row => {
      //     Object.keys(row).forEach(key => {
      //         row[key] = typeof row[key] === 'string' ? removerAcentos(row[key]) : row[key];
      //     });
      // });

      // Adiciona os dados ao array combinado
      combinedDadosData = combinedDadosData.concat(csvData);
      filesDadosProcessed++;

      // Se todos os arquivos foram processados, faça a manipulação dos dados
      if (filesDadosProcessed === filesDados.length) {
        Array.from(filesSip).forEach((file) => {
          const reader = new FileReader();

          reader.onload = function (event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // Assume que o CSV tem apenas uma folha (sheet)
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Converte o sheet para JSON
            const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 0 });

            // Adiciona os dados ao array combinado
            combinedSipData = combinedSipData.concat(csvData);
            filesSipProcessed++;

            // Se todos os arquivos foram processados, faça a manipulação dos dados
            if (filesSipProcessed === filesSip.length) {
              Array.from(filesVoz).forEach((file) => {
                const reader = new FileReader();

                reader.onload = function (event) {
                  const data = new Uint8Array(event.target.result);
                  const workbook = XLSX.read(data, { type: "array" });

                  // Assume que o CSV tem apenas uma folha (sheet)
                  const sheetName = workbook.SheetNames[0];
                  const worksheet = workbook.Sheets[sheetName];

                  // Converte o sheet para JSON
                  const csvData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 0,
                  });

                  // Adiciona os dados ao array combinado
                  combinedVozData = combinedVozData.concat(csvData);
                  filesVozProcessed++;

                  // Se todos os arquivos foram processados, faça a manipulação dos dados
                  if (filesVozProcessed === filesVoz.length) {
                    ManipulateDadosData(combinedDadosData);
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
// === === === Leitura dos arquivos === === ===

// === === === Processamento dos dados === === ===
function ManipulateDadosData(data) {
  console.log(data);

  data.forEach((line) => {
    let dataAssCompleta = numeroInteiroParaData(line["DT_ASS"]);
    let dataAssFormatada = numberToDate(dataAssCompleta);
    let dataFimCompleta = numeroInteiroParaData(
      line["DT_FIM_PRAZO_CONTRATUAL"]
    );
    let dataFimFormatada = numberToDate(dataFimCompleta);

    dataExport.push({
      CD_PESSOA: line["CD_PESSOA"],
      NMLINHANEGOCIO: "DADOS",
      NMLINHAPRODUTO: "PLANO FLEX",
      VELOCIDADE: line["VELOCIDADE"],
      NRCNPJ: line["NRCNPJ"],
      NUM_MESES_PRAZO_CONTRATUAL: line["NUM_MESES_PRAZO_CONTRATUAL"],
      DT_ASS: dataAssFormatada,
      DT_FIM_PRAZO_CONTRATUAL: dataFimFormatada,
      VALOR_TOTAL: Number(line["VALOR_TOTAL"]),
      END_CIDADE: 0,
      END_ESTADO: 0,
      END_CEP: line["END_CEP"],
      PERCENTUAL_DE_CONTRATO: line["CLIENTE"],
      TERMINAL_TRONCO: line["SISTEMA_ORIGEM"],
    });
  });

  ManipulateSipData(combinedSipData);
}

function ManipulateSipData(data) {
  console.log(data);

  data.forEach((line) => {
    let dataCompleta = line["DATA_CRIACAO"].slice(0, -13).split("-");
    let dia = dataCompleta[2];
    let mes = dataCompleta[1];
    let ano = parseInt(dataCompleta[0]) + 3;

    let valorFormatado = String(line["VL_TOTAL"]).slice(0, 3);

    if (String(valorFormatado).includes(".")) {
      if (String(valorFormatado).charAt(valorFormatado.length - 1) == ".") {
        valorFormatado = Number(String(valorFormatado).slice(0, -1)).toFixed(2);
      } else {
        valorFormatado = Number(valorFormatado).toFixed(2);
      }
    }

    dataExport.push({
      CD_PESSOA: line["COD_CLIENTE"],
      NMLINHANEGOCIO: "SIP",
      NMLINHAPRODUTO: "PLANO FLEX",
      VELOCIDADE: line["QTD_CANAIS"],
      NRCNPJ: line["CNPJ"],
      NUM_MESES_PRAZO_CONTRATUAL: 36,
      DT_ASS: line["DATA_CRIACAO"].slice(0, -13),
      DT_FIM_PRAZO_CONTRATUAL: `${ano}-${mes}-${dia}`,
      VALOR_TOTAL: valorFormatado,
      END_CIDADE: 0,
      END_ESTADO: 0,
      END_CEP: line["END_CEP"],
      PERCENTUAL_DE_CONTRATO: line["CLIENTE"],
      TERMINAL_TRONCO: line["TERMINAL"],
    });
  });

  ManipulateVozData(combinedVozData);
}

function ManipulateVozData(data) {
  console.log(data);

  data.forEach((line) => {
    let valorFormatado = String(line["VL_TOTAL"]).slice(0, 3);

    if (String(valorFormatado).includes(".")) {
      if (String(valorFormatado).charAt(valorFormatado.length - 1) == ".") {
        valorFormatado = Number(String(valorFormatado).slice(0, -1)).toFixed(2);
      } else {
        valorFormatado = Number(valorFormatado).toFixed(2);
      }
    }

    let cep = "";
    let match = "";
    let string = "";
    let produto = "";

    if (ufs.includes(String(line["UF"]))) {
      cep = line["NR_CEP"];
      produto = line["TP_PRODUTO"];

      string = String(line["TP_PRODUTO"]);
      if (string.includes(" ")) {
        [, match] = string.match(/(\S+) /) || [];
      } else {
        match = line["TP_PRODUTO"];
      }
    } else if (ufs.includes(String(line["TP_PRODUTO"]))) {
      cep = line["NR_IMOVEL"];
      produto = line["REDE"];

      string = String(line["REDE"]);
      if (string.includes(" ")) {
        [, match] = string.match(/(\S+) /) || [];
      } else {
        match = line["REDE"];
      }
    } else if (ufs.includes(String(line["REDE"]))) {
      cep = line["DS_BAIRRO"];
      produto = line["QTD_PRODUTOS_CONTA"];

      string = String(line["QTD_PRODUTOS_CONTA"]);
      if (string.includes(" ")) {
        [, match] = string.match(/(\S+) /) || [];
      } else {
        match = line["QTD_PRODUTOS_CONTA"];
      }
    } else {
      console.error(line["NR_CEP"]);
      cep = line["NR_CEP"];
      produto = line["TP_PRODUTO"];

      string = String(line["TP_PRODUTO"]);
      if (string.includes(" ")) {
        [, match] = string.match(/(\S+) /) || [];
      } else {
        match = line["TP_PRODUTO"];
      }
    }

    dataExport.push({
      CD_PESSOA: line["CD_PESSOA"],
      NMLINHANEGOCIO: match,
      NMLINHAPRODUTO: produto,
      VELOCIDADE: line["QTD_ACESSO"],
      NRCNPJ: "0",
      NUM_MESES_PRAZO_CONTRATUAL: 0,
      DT_ASS: "0",
      DT_FIM_PRAZO_CONTRATUAL: "0",
      VALOR_TOTAL: valorFormatado,
      END_CIDADE: "0",
      END_ESTADO: "0",
      END_CEP: cep,
      PERCENTUAL_DE_CONTRATO: "0",
      TERMINAL_TRONCO: line["TRONCO"],
    });
  });

  atualizarCep();
  console.warn("Terminou leitura");
}

function mostrarDadosProcessados(data) {
  console.log(data);
  console.warn("Dados informados");

  pProcess.setAttribute("style", "display: none");
  loader.setAttribute("style", "display: none");
  exportToExcel(dataExport, "Avancada", "Avancada.xlsx");
}
// === === === Processamento dos dados === === ===

// === === === Funcoes de atualizacao === === ===
function atualizarCep() {
  dataExport.forEach((line) => {
    let valor = line["END_CEP"];

    let maisProximo = faixaCep.reduce(function (anterior, corrente) {
      return Math.abs(corrente - valor) < Math.abs(anterior - valor)
        ? corrente
        : anterior;
    });

    if (line["END_CEP"] !== "END_CEP") {
      cepData.forEach((eB) => {
        if (maisProximo === eB.col3 || maisProximo === eB.col4) {
          line["END_ESTADO"] = eB.col1;
          line["END_CIDADE"] = eB.col2;
        }
      });
    }
  });
  dataDiff();
  console.log("Terminou Atualizar CEP");
}

function dataDiff() {
  dataExport.forEach((line) => {
    if (line["CD_PESSOA"] !== "CD_PESSOA") {
      let dataInicial = new Date(line["DT_ASS"]);
      let dataFinal = new Date(line["DT_FIM_PRAZO_CONTRATUAL"]);
      let hoje = new Date();

      if (line["DT_ASS"] !== "0") {
        let result1 = Math.abs(dataFinal - dataInicial);
        let days1 = result1 / (1000 * 3600 * 24);

        let result2 = Math.abs(hoje - dataInicial);
        let days2 = result2 / (1000 * 3600 * 24);

        let result3 = days2 / days1;
        line["PERCENTUAL_DE_CONTRATO"] = result3.toFixed(3);
      } else {
        line["PERCENTUAL_DE_CONTRATO"] = 1;
      }
    }
  });
  verificacaoFinal();
  console.log("Terminou Percentual");
}

function verificacaoFinal() {
  dataExport.forEach((line) => {
    line["PERCENTUAL_DE_CONTRATO"] = Number(line["PERCENTUAL_DE_CONTRATO"]);
    line["VALOR_TOTAL"] = Number(line["VALOR_TOTAL"]);

    line["DT_ASS"] = new Date(line["DT_ASS"]);
    let dataBase2 = new Date(1900, 0, -2);
    let dataInt = parseInt(
      (line["DT_ASS"].getTime() - dataBase2) / 24 / 60 / 60 / 1000
    );
    line["DT_ASS"] = dataInt;
    line["DT_FIM_PRAZO_CONTRATUAL"] = new Date(line["DT_FIM_PRAZO_CONTRATUAL"]);
    let dataMovimentacao = parseInt(
      (line["DT_FIM_PRAZO_CONTRATUAL"].getTime() - dataBase2) /
        24 /
        60 /
        60 /
        1000
    );
    line["DT_FIM_PRAZO_CONTRATUAL"] = dataMovimentacao;

    if (
      line["END_CEP"] == undefined ||
      line["END_CEP"] == 0 ||
      line["END_CEP"] == "" ||
      line["END_CEP"] == 1
    ) {
      line["END_CEP"] = 0;
      line["END_ESTADO"] = 0;
      line["END_CIDADE"] = 0;
    }
  });

  mostrarDadosProcessados(dataExport);
}
// === === === Funcoes de atualizacao === === ===
