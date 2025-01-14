const inputMovel = document.getElementById("fileInputMovel");
const inputFixa = document.getElementById("fileInputFixa");
const inputAvancada = document.getElementById("fileInputAvancada");

const labelMovel = document.getElementById("labelFileInputMovel");
const labelFixa = document.getElementById("labelFileInputFixa");
const labelAvancada = document.getElementById("labelFileInputAvancada");

let combinedMovelData = [];
let combinedFixaData = [];
let combinedAvancadaData = [];

let filesMovelProcessed = 0;
let filesFixaProcessed = 0;
let filesAvancadaProcessed = 0;

const dataExport = [];

// === === === Modificacao label === === ===
inputMovel.addEventListener("change", (event) => {
  labelMovel.textContent = event.target.files[0].name;
});
inputFixa.addEventListener("change", (event) => {
  labelFixa.textContent = event.target.files[0].name;
});
inputAvancada.addEventListener("change", (event) => {
  labelAvancada.textContent = event.target.files[0].name;
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

function separarPorChave(arr, chave) {
  // funcao para separar um array de objetos pela chave
  let obj = {};
  arr.forEach((item) => {
    let key = item[chave];
    if (!obj[key]) {
      obj[key] = [];
    }
    obj[key].push(item);
  });
  console.log(obj);
  return obj;
}
// === === === Funcoes auxiliares === === ===

// === === === Leitura dos arquivos === === ===
document.getElementById("processBtn").addEventListener("click", function () {
  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const filesMovel = inputMovel.files;
  const filesFixa = inputFixa.files;
  const filesAvancada = inputAvancada.files;

  if (
    filesMovel.length === 0 ||
    filesFixa.length === 0 ||
    filesAvancada.length === 0
  ) {
    alert("Por favor, selecione pelo menos um arquivo para cada input.");
    return;
  }

  Array.from(filesMovel).forEach((file) => {
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
      combinedMovelData = combinedMovelData.concat(csvData);
      filesMovelProcessed++;

      // Se todos os arquivos foram processados, faça a manipulação dos dados
      if (filesMovelProcessed === filesMovel.length) {
        Array.from(filesFixa).forEach((file) => {
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
            combinedFixaData = combinedFixaData.concat(csvData);
            filesFixaProcessed++;

            // Se todos os arquivos foram processados, faça a manipulação dos dados
            if (filesFixaProcessed === filesFixa.length) {
              Array.from(filesAvancada).forEach((file) => {
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
                  combinedAvancadaData = combinedAvancadaData.concat(csvData);
                  filesAvancadaProcessed++;

                  // Se todos os arquivos foram processados, faça a manipulação dos dados
                  if (filesAvancadaProcessed === filesAvancada.length) {
                    ManipulateMovelData(combinedMovelData);
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

// === === === Manipulacao dos dados === === ===
function ManipulateMovelData(data) {
  console.log(data);
  let objetoSeparado = separarPorChave(data, "");
  console.log(objetoSeparado);

  ManipulateFixaData(combinedFixaData);
}

function ManipulateFixaData(data) {
  console.log(data);

  ManipulateAvancadaData(combinedAvancadaData);
}

function ManipulateAvancadaData(data) {
  console.log(data);

  loader.setAttribute("style", "display: none");
  pProcess.setAttribute("style", "display: none");
  // exportToExcel(dataExport, "Avancada", "Avancada.xlsx");
}
// === === === Manipulacao dos dados === === ===
