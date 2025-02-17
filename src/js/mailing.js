const pProcess = document.getElementById("pProcessando");
const loader = document.getElementById("loader");

const labelCall = document.getElementById("labelCallFileInput");
const labelDisc = document.getElementById("labelDiscFileInput");
const labelInput = document.getElementById("labelInputFileInput");

let csvData = [];
let callData = [];
let inputData = [];

document
  .getElementById("fileInputInput")
  .addEventListener("change", handleInput, false);

document
  .getElementById("fileInput")
  .addEventListener("change", handleFileSelect, false);

document
  .getElementById("fileCallInput")
  .addEventListener("change", handleCall, false);

const mailing = [];

function handleInput() {
  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const filesInput = document.getElementById("fileInputInput").files;

  if (filesInput.length > 1) {
    labelInput.textContent = `${filesInput.length} arquivos`;
  } else {
    labelInput.textContent = `${filesInput.length} arquivo`;
  }

  Array.from(filesInput).forEach((file) => {
    const reader = new FileReader();

    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      inputData = inputData.concat(csvData);
      console.log(inputData);
    };

    pProcess.setAttribute("style", "display: none");
    loader.setAttribute("style", "display: none");

    reader.readAsArrayBuffer(file);
  });
}

function handleCall(event) {
  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const file = event.target.files[0];
  labelCall.textContent = file.name;

  if (!file) {
    alert("Por favor, selecione um arquivo.");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet);

    // Dividir os dados pela coluna C
    const splitFiles = {};

    json.forEach((row) => {
      const key = row["ID_Cliente"];
      if (!splitFiles[key]) {
        splitFiles[key] = [];
      }
      splitFiles[key].push(row);
    });

    const splitFileContents = [];

    Object.keys(splitFiles).forEach((key) => {
      const newWorksheet = XLSX.utils.json_to_sheet(splitFiles[key]);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
      const csvContent = XLSX.write(newWorkbook, {
        bookType: "csv",
        type: "string",
      });

      // Armazenar o conteúdo no array
      splitFileContents.push({ fileName: `${key}.csv`, content: csvContent });
    });

    splitFileContents.forEach((file) => {
      // console.log(`Processing file: ${file.fileName}`);
      const rows = XLSX.utils.sheet_to_json(
        XLSX.read(file.content, { type: "string" }).Sheets.Sheet1
      );

      rows.forEach((e) => {
        ttt = parseInt(e.Tempo_total * 86400);

        callData.push({
          n: e.NÃºmero,
          t: ttt,
        });
      });
    });

    pProcess.setAttribute("style", "display: none");
    loader.setAttribute("style", "display: none");
  };

  reader.readAsArrayBuffer(file);
}

function handleFileSelect(event) {
  loader.setAttribute("style", "display: block");
  pProcess.setAttribute("style", "display: block");

  const file = event.target.files[0];
  labelDisc.textContent = file.name;

  if (!file) {
    alert("Por favor, selecione um arquivo.");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet);

    const splitFiles = {};

    json.forEach((row) => {
      const key = row["ID_Cliente"];
      if (!splitFiles[key]) {
        splitFiles[key] = [];
      }
      splitFiles[key].push(row);
    });

    const splitFileContents = [];

    Object.keys(splitFiles).forEach((key) => {
      const newWorksheet = XLSX.utils.json_to_sheet(splitFiles[key]);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
      const csvContent = XLSX.write(newWorkbook, {
        bookType: "csv",
        type: "string",
      });

      splitFileContents.push({ fileName: `${key}`, content: csvContent });
    });

    splitFileContents.forEach((file) => {
      // console.log(`Processing file: ${file.fileName}`);
      const rows = XLSX.utils.sheet_to_json(
        XLSX.read(file.content, { type: "string" }).Sheets.Sheet1
      );

      let tipo = "";
      let dataSemBarra;

      rows.forEach((e) => {
        let data;
        let dia;
        let mes;
        let ano;
        let dataCompleta;

        if (String(e.Data).includes("/")) {
          dia = String(e.Data).substring(0, 2);
          mes = String(e.Data).substring(3, 5);
          ano = String(e.Data).substring(6, 10);
          dataCompleta = `${ano}-${mes}-${dia}`;
          dataSemBarra = `${ano}${mes}${dia}`;
        } else {
          let numero = e.Data + 1;
          data = numeroInteiroParaData(numero);

          dia = data.getUTCDate() - 1;
          mes = data.getUTCMonth() + 1;
          ano = data.getUTCFullYear();

          if (String(dia).length == 1) {
            dia = `0${dia}`;
          }

          if (String(mes).length == 1) {
            mes = `0${mes}`;
          }

          dataCompleta = `${ano}-${dia}-${mes}`;
          dataSemBarra = `${ano}${dia}${mes}`;
        }

        let totalSeconds = e.Hora * 24 * 3600;

        let horas = String(Math.floor(totalSeconds / 3600)).padStart(2, "0");
        let minutos = String(Math.floor((totalSeconds % 3600) / 60)).padStart(
          2,
          "0"
        );
        let segundos = String(Math.floor(totalSeconds % 60)).padStart(2, "0");
        let horario = `${horas}:${minutos}:${segundos}`;

        let dataEHora = `${dataCompleta} ${horario}`;

        let ddd = e.DDD;
        let telefone = e.Telefone;

        let telefoneCompleto = `${ddd}${telefone}`;

        let audiencia = "";
        let tipoAudiencia = 0;
        let terminais = 0;

        if (e["ISDN_Code"] == 29 || e["ISDN_Code"] == 58) {
          // console.warn('AMD ou Falha da Operadora')
        } else {
          if (e.Campanha.includes("MIGRACAO")) {
            audiencia = telefoneCompleto;
            tipoAudiencia = 3;
          } else {
            if (String(e.Field_2).startsWith("'")) {
              audiencia = e.Field_2.substring(1);
            } else {
              audiencia = e.Field_2;
            }
            tipoAudiencia = 4;

            let index = inputData.findIndex(
              (element) => element[20] == telefoneCompleto
            );

            if (index != -1) {
              if (String(inputData[index][17]).includes("TOTAL")) {
                terminais = 2;
              } else if (String(inputData[index][17]).includes("SOLO")) {
                terminais = 1;
              }
            }
          }

          if (tipoAudiencia == 3) {
            tipo = "MOV";
          } else if (tipoAudiencia == 4) {
            tipo = "FIX";
          }

          if (tipoAudiencia == 3) {
            csvData.push({
              ID_PLAY: e.ID_Cliente,
              ID_AUDIENCIA: audiencia,
              ID_TIPO_AUDIENCIA: tipoAudiencia,
              ID_MOTIVO_TABULACAO: e.ISDN_Code,
              DT_TABULACAO: dataEHora,
              NR_TLFN: telefoneCompleto,
              NOVO_CPF: null,
              CPF_OPERADOR: null,
              NR_TLFN_ADD1: null,
              NR_TLFN_ADD2: null,
              NR_TLFN_ADD3: null,
              EMAIL_ADD1: null,
              NR_DURACAO_CHAMADA: null,
              QTD_LIGACAO: 1,
              CORINGA1: null,
              CORINGA2: null,
            });
          } else if (tipoAudiencia == 4) {
            csvData.push({
              ID_CAMPANHA: e.ID_Cliente,
              ID_AUDIENCIA: audiencia,
              ID_TIPO_AUDIENCIA: tipoAudiencia,
              ID_RETORNO: e.Disposition_Return_Code,
              DT_TABULACAO: dataEHora,
              NR_TLFN: telefoneCompleto,
              NOVO_CPF: null,
              CPF_OPERADOR: null,
              NR_TLFN_ADD1: null,
              NR_TLFN_ADD2: null,
              NR_TLFN_ADD3: null,
              EMAIL_ADD1: null,
              NR_DURACAO_CHAMADA: null,
              QT_CHAMADAS: 1,
              ID_OFERTA: null,
              QT_TERMINAIS: terminais,
            });
          }
        }
      });

      csvData.forEach((e) => {
        let index = callData.findIndex((element) => element.n == e.NR_TLFN);
        if (index != -1) {
          e.NR_DURACAO_CHAMADA = String(callData[index].t);
        }
      });

      let secondsNow = Date.now();
      let dataHoje = new Date(secondsNow);

      let horasHoje = dataHoje.getHours();
      let minutosHoje = dataHoje.getMinutes();
      let segundosHoje = dataHoje.getSeconds();

      if (String(horasHoje).length == 1) {
        horasHoje = `0${horasHoje}`;
      }

      if (String(minutosHoje).length == 1) {
        minutosHoje = `0${minutosHoje}`;
      }

      if (String(segundosHoje).length == 1) {
        segundosHoje = `0${segundosHoje}`;
      }

      let horarioHoje = `${horasHoje}${minutosHoje}${segundosHoje}`;

      console.log(csvData);
      if (csvData.length == 0) {
        console.info("Nenhum dado encontrado.");
        return;
      }

      const headers = Object.keys(csvData[0]).join("|");
      const txtContent = csvData
        .map((obj) => Object.values(obj).join("|"))
        .join("\r\n");
      const fullContent = `${headers}\r\n${txtContent}`;
      // Cria o blob e baixa o arquivo
      const blob = new Blob([fullContent], {
        type: "text/plain;charset=utf-8;",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${file.fileName}_RET_${tipo}_MS_${dataSemBarra}${horarioHoje}.txt`; // Altere o nome para .txt
      a.click();
      // downloadManipulatedCsv(csvData, `${file.fileName}_RET_${tipo}_MS_${dataSemBarra}${horarioHoje}.csv`)
      csvData = [];
    });

    pProcess.setAttribute("style", "display: none");
    loader.setAttribute("style", "display: none");
  };

  reader.readAsArrayBuffer(file);
}

function numeroInteiroParaData(numero) {
  var dataBase = new Date(1900, 0, -1); // Data base do Excel (1º de janeiro de 1900)
  var data = new Date(dataBase.getTime() + numero * 24 * 60 * 60 * 1000);
  return data;
}

function downloadManipulatedCsv(data, name) {
  const newWorksheet = XLSX.utils.json_to_sheet(data);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
  const csvContent = XLSX.write(newWorkbook, {
    bookType: "csv",
    type: "string",
    FS: "|",
  });

  downloadCSV(csvContent, name);
}

function downloadCSV(csvContent, fileName) {
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");

  if (link.download !== undefined) {
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", fileName);
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
}
