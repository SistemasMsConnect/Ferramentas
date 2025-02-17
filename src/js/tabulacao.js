const pProcess = document.getElementById("pProcessando");
const loader = document.getElementById("loader");

const inputFixa = document.getElementById("fileFixaInput");
const inputAtivos = document.getElementById("fileAtivosInput");
const dropdown = document.getElementById("dropdown");

const labelFixa = document.getElementById("labelFixaFileInput");
const labelAtivos = document.getElementById("labelAtivosFileInput");

let dataFixaExport = [];

let filesFixaProcessed = 0;
let filesAtivosProcessed = 0;

let combinedFixaData = [];
let combinedAtivosData = [];

inputFixa.addEventListener("change", (event) => {
  if (event.target.files.length > 1) {
    labelFixa.textContent = `${event.target.files.length} arquivos`;
  } else if (event.target.files.length == 1) {
    labelFixa.textContent = `${event.target.files.length} arquivo`;
  }
});

inputAtivos.addEventListener("change", (event) => {
  if (event.target.files.length > 1) {
    labelAtivos.textContent = `${event.target.files.length} arquivos`;
  } else if (event.target.files.length == 1) {
    labelAtivos.textContent = `${event.target.files.length} arquivo`;
  }
});

document.getElementById("processBtn").addEventListener("click", function () {
  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const filesFixa = inputFixa.files;
  const filesAtivos = inputAtivos.files;

  if (
    filesFixa.length === 0 ||
    filesAtivos.length === 0
  ) {
    alert("Por favor, selecione pelo menos um arquivo CSV.");
    return;
  }

  Array.from(filesFixa).forEach((file) => {
    const reader = new FileReader();

    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });

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
        Array.from(filesAtivos).forEach((file) => {
          const reader = new FileReader();

          reader.onload = function (event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // Assume que o CSV tem apenas uma folha (sheet)
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Converte o sheet para JSON
            const csvData = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
            });

            // Adiciona os dados ao array combinado
            combinedAtivosData = combinedAtivosData.concat(csvData);

            manipulateFixaData(combinedFixaData);
          };

          reader.readAsArrayBuffer(file);
        });
      }
    };

    reader.readAsArrayBuffer(file);
  });
});

function manipulateFixaData(data) {
  console.log(data)
  data.forEach((e) => {
    if (e[8] == "AlÃ´" || e[8] == "Alo") {
      let funcional = "";
      let index = combinedAtivosData.findIndex(
        (element) =>
          element[0] == e[14] ||
          String(element[5]).replace(/[^0-9]/g, "") == e[15]
      );
      if (index != -1) {
        funcional = combinedAtivosData[index][7];
      }

      let dataTabulacao = "";

      if (typeof e[3] == "number") {
        const parsedDate = XLSX.SSF.parse_date_code(Number(e[3]));
        const year = parsedDate.y;
        const month = String(parsedDate.m).padStart(2, "0");
        const day = String(parsedDate.d).padStart(2, "0");
        dataTabulacao = `${month}/${day}/${year}`;
      } else {
        dataTabulacao = e[3].split(" ")[0];
      }

      dataFixaExport.push({
        TERMINAL: `${e[5]}${e[6]}`,
        NOME_CLIENTE: e[15],
        CPF_CLIENTE: e[17],
        FUNCIONAL: funcional,
        OFERTA_DISPONIVEL: e[25],
        PRODUTO: "Fixa",
        TABULACAO: e[9],
        DATA: dataTabulacao,
        SEMANA: dropdown.value,
        ID_PLAY: e[14],
      });
    }
  });

  pProcess.setAttribute("style", "display: none");
  loader.setAttribute("style", "display: none");

  exportToCSV(
    dataFixaExport,
    "Tabulacao Fixa",
    "dataTabulacao.xlsx"
  );
}

function exportToCSV(
  dataFixa,
  sheetnameFixa,
  filename
) {
  const worksheet = XLSX.utils.json_to_sheet(dataFixa);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, `${sheetnameFixa}`);

  XLSX.writeFile(workbook, filename, { bookType: "xlsx" });
}
