const pProcess = document.getElementById("pProcessando");
const loader = document.getElementById("loader");

const inputFixa = document.getElementById("fileFixaInput");

const labelFixa = document.getElementById("labelFixaFileInput");

let dataFixaExport = [];

let filesFixaProcessed = 0;

let combinedFixaData = [];

inputFixa.addEventListener("change", (event) => {
  if (event.target.files.length > 1) {
    labelFixa.textContent = `${event.target.files.length} arquivos`;
  } else if (event.target.files.length == 1) {
    labelFixa.textContent = `${event.target.files.length} arquivo`;
  }
});

document.getElementById("processBtn").addEventListener("click", function () {
  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const filesFixa = inputFixa.files;

  if (filesFixa.length === 0) {
    alert("Por favor, selecione pelo menos um arquivo CSV.");
    return;
  }

  Array.from(filesFixa).forEach((file) => {
    const reader = new FileReader();

    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      combinedFixaData = combinedFixaData.concat(csvData);
      filesFixaProcessed++;

      if (filesFixaProcessed === filesFixa.length) {
        manipulateFixaData(combinedFixaData)
      }
    };

    reader.readAsArrayBuffer(file);
  });
});

function manipulateFixaData(data) {
  data.forEach((e) => {
    if (e[0] != "TERMINAL") {
      dataFixaExport.push({
        TERMINAL: e[0],
        NOME_CLIENTE: e[1],
        CPF_CLIENTE: e[2],
        FUNCIONAL: e[3],
        OFERTA_DISPONIVEL: e[4],
        PRODUTO: e[5],
        TABULACAO: e[6],
        DATA: e[7],
        SEMANA: e[8],
        ID_PLAY: e[9]
      })
    }
  })

  loader.setAttribute("style", "display: none");
  pProcess.setAttribute("style", "display: none");

  exportToCSV(dataFixaExport, "data", "data.xlsx")
}

function exportToCSV(dataFixa, sheetnameFixa, filename) {
  const worksheet = XLSX.utils.json_to_sheet(dataFixa);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, `${sheetnameFixa}`);

  XLSX.writeFile(workbook, filename, { bookType: "xlsx" });
}
