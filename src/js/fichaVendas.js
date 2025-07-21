const pProcess = document.getElementById("pProcessando");
const loader = document.getElementById("loader");

const inputInput = document.getElementById("fileInputInput");
const inputTabulacao = document.getElementById("fileTabulacaoInput");
const inputOitenta = document.getElementById("fileOitentaInput");

const labelInput = document.getElementById("labelInputFileInput");
const labelTabulacao = document.getElementById("labelTabulacaoFileInput");
const labelOitenta = document.getElementById("labelOitentaFileInput");

let dataMovelExport = [];
let dataFixaExport = [];
let dataAceitesExport = [];

let filesInputProcessed = 0;
let filesTabulacaoProcessed = 0;
let filesOitentaProcessed = 0;

let combinedInputData = [];
let combinedTabulacaoData = [];
let combinedOitentaData = [];

inputInput.addEventListener("change", (event) => {
  if (event.target.files.length > 1) {
    labelInput.textContent = `${event.target.files.length} arquivos`;
  } else if (event.target.files.length == 1) {
    labelInput.textContent = `${event.target.files.length} arquivo`;
  }
});

inputTabulacao.addEventListener("change", (event) => {
  if (event.target.files.length > 1) {
    labelTabulacao.textContent = `${event.target.files.length} arquivos`;
  } else if (event.target.files.length == 1) {
    labelTabulacao.textContent = `${event.target.files.length} arquivo`;
  }
});

inputOitenta.addEventListener("change", (event) => {
  if (event.target.files.length > 1) {
    labelOitenta.textContent = `${event.target.files.length} arquivos`;
  } else if (event.target.files.length == 1) {
    labelOitenta.textContent = `${event.target.files.length} arquivo`;
  }
});

document.getElementById("processBtn").addEventListener("click", function () {
  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const filesInput = inputInput.files;
  const filesTabulacao = inputTabulacao.files;
  const filesOitenta = inputOitenta.files;

  if (
    filesInput.length === 0 ||
    filesTabulacao.length === 0 ||
    filesOitenta.length === 0
  ) {
    alert("Por favor, selecione pelo menos um arquivo em cada campo.");
    return;
  }

  Array.from(filesInput).forEach((file) => {
    const reader = new FileReader();

    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Assume que o CSV tem apenas uma folha (sheet)
      const sheetName = workbook.SheetNames[ 0 ];
      const worksheet = workbook.Sheets[ sheetName ];

      // Converte o sheet para JSON
      const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Adiciona os dados ao array combinado
      combinedInputData = combinedInputData.concat(csvData);
      filesInputProcessed++;

      // Se todos os arquivos foram processados, faça a manipulação dos dados
      if (filesInputProcessed === filesInput.length) {
        Array.from(filesTabulacao).forEach((file) => {
          const reader = new FileReader();

          reader.onload = function (event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // Assume que o CSV tem apenas uma folha (sheet)
            const sheetName = workbook.SheetNames[ 0 ];
            const worksheet = workbook.Sheets[ sheetName ];

            // Converte o sheet para JSON
            const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Adiciona os dados ao array combinado
            combinedTabulacaoData = combinedTabulacaoData.concat(csvData);
            filesTabulacaoProcessed++;

            // Se todos os arquivos foram processados, faça a manipulação dos dados
            if (filesTabulacaoProcessed === filesTabulacao.length) {
              Array.from(filesOitenta).forEach((file) => {
                const reader = new FileReader();

                reader.onload = function (event) {
                  const data = new Uint8Array(event.target.result);
                  const workbook = XLSX.read(data, { type: "array" });

                  // Assume que o CSV tem apenas uma folha (sheet)
                  const sheetName = workbook.SheetNames[ 0 ];
                  const worksheet = workbook.Sheets[ sheetName ];

                  // Converte o sheet para JSON
                  const csvData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                  });

                  // Adiciona os dados ao array combinado
                  combinedOitentaData = combinedOitentaData.concat(csvData);
                  filesOitentaProcessed++;

                  // Se todos os arquivos foram processados, faça a manipulação dos dados
                  if (filesOitentaProcessed === filesOitenta.length) {
                    manipulateTabulacaoData(combinedTabulacaoData);
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

function manipulateTabulacaoData(data) {
  console.log(data);

  let aceitesMovel = 0;
  let aceitesFixa = 0;

  data.forEach((e) => {
    if (e[ 9 ] == "VENDA") {
      if (e[ 1 ].includes("FIXA")) {
        aceitesFixa++;
      } else if (e[ 1 ].includes("MIGRACAO")) {
        aceitesMovel++;
      }
    }
  });

  dataAceitesExport.push({
    aceitesMovel: aceitesMovel,
    aceitesFixa: aceitesFixa,
  });

  manipulateInputData(combinedInputData);
}

function manipulateInputData(data) {
  console.log(data);

  data.forEach((e) => {
    let campanha = "";
    let dataCompletaTabulacao = "";
    let regiao = "";
    let adicional = "";
    let cod = "";
    let oferta = "";
    let status = e[ 48 ];
    let subStatus = e[ 49 ];
    let trocaTitularidade = e[ 47 ];
    let oitenta = "";
    let aceite = e[ 25 ].includes('DÉBITO') ? 'SIM' : 'NÃO'
    let words = ''
    if (e[ 10 ] != undefined) {
      words = e[ 10 ].split(' ')
      for (let i = 0; i < words.length; i++) {
        if (words[ i ][ 0 ] !== undefined) {
          words[ i ] = words[ i ][ 0 ].toUpperCase() + words[ i ].substr(1).toLowerCase()
        }
      }
    }
    let cidade = '';

    let index80 = combinedOitentaData.findIndex(element => element[ 6 ] == e[ 37 ])

    if (index80 !== -1) {
      oitenta = combinedOitentaData[ index80 ][ 5 ];
    }

    if (String(e[ 17 ]).includes("|")) {
      oferta = String(e[ 17 ]).split("|")[ 0 ];
    } else {
      oferta = e[ 17 ];
    }

    if (String(e[ 29 ]).toLowerCase().includes("fixa")) {
      campanha = "FIXA FTTH";
    } else if (String(e[ 29 ]).toLowerCase().includes("mig")) {
      campanha = "MIGRAÇÃO PRE CTRL";
    }

    let indexTabulacao = combinedTabulacaoData.findIndex(
      (element) =>
        `${element[ 5 ]}${element[ 6 ]}` == e[ 7 ] ||
        `${element[ 5 ]}${element[ 6 ]}` == e[ 8 ] ||
        (String(element[ 17 ]).replace(/[^0-9]/g, "") == String(e[ 4 ]).replace(/[^0-9]/g, "")) &&
        element[ 10 ] == "VENDA"
    );
    if (indexTabulacao != -1) {
      if (!String(combinedTabulacaoData[ indexTabulacao ][ 1 ]).includes("fixa")) {
        adicional = `${combinedTabulacaoData[ indexTabulacao ][ 30 ]} + 5`;
      }

      if (String(combinedTabulacaoData[ indexTabulacao ][ 3 ]).length == 5) {
        let dataFuncao = numeroInteiroParaData(
          combinedTabulacaoData[ indexTabulacao ][ 3 ]
        );

        dataCompletaTabulacao = `${dataFuncao.getUTCMonth() + 1
          }/${dataFuncao.getUTCDate()}/${dataFuncao.getUTCFullYear()}`;
      } else {
        dataCompletaTabulacao = combinedTabulacaoData[ indexTabulacao ][ 3 ];
      }

      cidade = combinedTabulacaoData[ indexTabulacao ][ 32 ];
    }

    if ([ "11", "12", "13" ].includes(String(e[ 7 ]).slice(0, 2))) {
      regiao = "SP1";
    } else if (
      [ "14", "15", "16", "17", "18", "19" ].includes(String(e[ 7 ]).slice(0, 2))
    ) {
      regiao = "SPI";
    }

    if (String(e[ 16 ]).includes("MOVEL") || e[ 16 ] == "Dados da venda") {
      adicional = adicional;
    } else {
      adicional = "";
    }

    if (e[ 25 ] == "FATURA DIGITAL") {
      cod = "SIM";
    } else {
      cod = "NÃO";
    }

    let dataEmissao = "";

    if (typeof e[ 33 ] == "number") {
      const parsedDate = XLSX.SSF.parse_date_code(Number(e[ 33 ]));
      const year = parsedDate.y;
      const month = String(parsedDate.m).padStart(2, "0");
      const day = String(parsedDate.d).padStart(2, "0");
      dataEmissao = `${month}/${day}/${year}`;
    } else {
      if (e[ 33 ] == undefined || e[ 33 ] == null) {
        dataEmissao = '';
      } else {
        dataEmissao = e[ 33 ].split(" ")[ 0 ];
      }
    }

    // Modificado para o arquivo do modulo 89

    // if (campanha == "MIGRAÇÃO PRE CTRL") {
    //   dataMovelExport.push({
    //     Campanha: campanha,
    //     DataVenda: dataCompletaTabulacao,
    //     DataEmissao: dataEmissao,
    //     Eps: "MS CONNECT",
    //     Regional: regiao,
    //     Terminal: e[ 7 ],
    //     CPFCliente: String(e[ 4 ]).replace(/[^0-9]/g, ""),
    //     Matricula: oitenta,
    //     Oferta: oferta,
    //     Status: status,
    //     SubStatus: subStatus,
    //     TrocaTitularidade: trocaTitularidade,
    //     Mailing: e[ 0 ],
    //     Perfil: "",
    //     Site: "MS",
    //     PlataformaEmissao: "NEXT",
    //     HomeOffice: "NÃO",
    //     Cod: cod,
    //     EmailFatura: e[ 23 ],
    //     PacoteAdicional: adicional,
    //   });
    // } else
    if (campanha == "FIXA FTTH") {
      dataFixaExport.push({
        Campanha: campanha,
        DataVenda: dataCompletaTabulacao,
        DataEmissao: dataEmissao,
        Eps: "MS CONNECT",
        Regional: regiao,
        Terminal: e[ 7 ],
        CpfCliente: e[ 4 ],
        Funcional: oitenta,
        Oferta: oferta,
        ValorOferta: e[ 18 ],
        Status: status,
        SubStatus: subStatus,
        TrocaTitularidade: trocaTitularidade,
        Mailing: e[ 0 ],
        Perfil: "",
        Site: "MS",
        Plataforma: "NEXT",
        HomeOffice: "NÃO",
        Cod: cod,
        EmailFatura: e[ 23 ],
        PacoteAdicional1: e[ 21 ],
        ValorAdicional1: '',
        PacoteAdicional2: '',
        ValorAdicional2: '',
        PacoteAdicional3: '',
        ValorAdicional3: '',
        PacoteAdicional4: '',
        ValorAdicional4: '',
        PacoteAdicional5: '',
        ValorAdicional5: '',
        PacoteAdicional6: '',
        ValorAdicional6: '',
        Logradouro: words.join(' ') != undefined ? words.join(' ') : '',
        Numero: e[ 12 ],
        Complemento: e[ 11 ],
        Bairro: e[ 14 ],
        Cidade: cidade,
        UF: '',
        NumeroOS: e[ 30 ],
        AceiteDebitoEmConta: aceite,
        CodigoBanco: '',
        NomeBanco: e[ 26 ],
        AgenciaBanco: e[ 27 ],
        NumeroConta: e[ 28 ],
        DigitoConta: '',
        TitularConta: e[ 3 ]
      });
    }
  });

  exportToCSV(
    dataMovelExport,
    dataFixaExport,
    dataAceitesExport,
    "dataFichaVendas.xlsx"
  );

  loader.setAttribute("style", "display: none");
  pProcess.setAttribute("style", "display: none");
}

function numeroInteiroParaData(numero) {
  var dataBase = new Date(1900, 0, -1); // Data base do Excel (1º de janeiro de 1900)
  var data = new Date(dataBase.getTime() + numero * 24 * 60 * 60 * 1000);
  return data;
}

function exportToCSV(dataMovel, dataFixa, dataAceites, filename) {
  // Cria uma nova worksheet a partir dos dados filtrados
  // const worksheet1 = XLSX.utils.json_to_sheet(dataMovel);
  const worksheet2 = XLSX.utils.json_to_sheet(dataFixa);
  const worksheet3 = XLSX.utils.json_to_sheet(dataAceites);

  // Cria um novo workbook e adiciona a worksheet a ele
  const workbook = XLSX.utils.book_new();
  // XLSX.utils.book_append_sheet(workbook, worksheet1, "Movel");
  XLSX.utils.book_append_sheet(workbook, worksheet2, "Fixa");
  XLSX.utils.book_append_sheet(workbook, worksheet3, "Aceites");

  // Exporta o workbook como um arquivo XLSX
  XLSX.writeFile(workbook, filename, { bookType: "xlsx" });
}
