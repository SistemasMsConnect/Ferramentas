const loader = document.getElementById("loader");
const btn = document.getElementById("btn");
const fileName = document.getElementById("labelFileInput");
const pProcess = document.getElementById("pProcessando");
const input = document.getElementById("fileInput");

async function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Lendo apenas a primeira planilha
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "0" });

      resolve({ fileName: file.name, data: jsonData });
    };

    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

async function handleFiles(file) {
  console.log("Iniciando processamento..."); // Verifique se aparece 2x

  const { fileName, data } = await readExcelFile(file);
  console.log("Arquivo processado:", fileName, data); // Verifique se aparece 2x
}

document
  .getElementById("fileInput")
  .addEventListener("change", function (event) {
    loader.setAttribute("style", "display: block");
    pProcess.setAttribute("style", "display: block");

    var files = input.files;

    Array.from(files).forEach((file) => {
      handleFiles(file);
      //   const reader = new FileReader();

      //   reader.onload = function (event) {
      //     const data = new Uint8Array(event.target.result);
      //     const workbook = XLSX.read(data, { type: "array" });

      //     // Assume que o CSV tem apenas uma folha (sheet)
      //     const sheetName = workbook.SheetNames[0];
      //     const worksheet = workbook.Sheets[sheetName];

      //     // Converte o sheet para JSON
      //     const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      //     console.log(csvData);

      //     const headers = csvData[0];
      //     const rows = csvData.slice(1);

      //     const indexedData = rows.map((row) => {
      //       let obj = {};
      //       row.forEach((value, index) => {
      //         obj[headers[index]] = value;
      //       });
      //       return obj;
      //     });

      //     const content = indexedData.map((item) => {
      //       item.REDE = item.CONTA_COBRANCA;
      //       delete item.CONTA_COBRANCA;
      //       return item;
      //     });

      //     const colunasDesejadas = [
      //       "DOCUMENTO",
      //       "DESIGNADOR",
      //       "VL_FAT_BRUTO",
      //       "DS_PRODUTO",
      //       "REDE",
      //       "DT_INI_FDLZ",
      //       "DT_FIM_FDLZ",
      //       "TECNOLOGIA_ORIGEM",
      //       "DEPARA",
      //       "M",
      //       "DS_PRODUTO_TECNOLOGIA",
      //     ];

      //     const filteredData = content.map((item) => {
      //       return colunasDesejadas.reduce((obj, key) => {
      //         if (item[key] !== undefined) {
      //           obj[key] = item[key];
      //         }
      //         return obj;
      //       }, {});
      //     });

      //     console.log(filteredData);
      //     processCSV(filteredData);
      //   };

      //   reader.readAsArrayBuffer(file);
    });
  });

function processCSV(content) {
  content.forEach(function (line) {
    if (line["DEPARA"] != undefined) {
      line["DEPARA"] = line["DEPARA"].trim();
    }

    if (
      line["DEPARA"] === "" ||
      line["DEPARA"] === null ||
      line["DEPARA"] === undefined
    ) {
      line["DEPARA"] = "zzz";
    }

    if (line["M"] === "" || line["M"] === null || line["M"] === undefined) {
      line["M"] = "0";
    }

    if (
      line["DOCUMENTO"] === "" ||
      line["DOCUMENTO"] === null ||
      line["DOCUMENTO"] === undefined
    ) {
      line["DOCUMENTO"] = "0";
    }
  });

  content.sort((a, b) => {
    return a["DEPARA"].localeCompare(b["DEPARA"]);
  });

  content.forEach((e, index) => {
    if (e["DEPARA"] === "zzz") {
      e["DEPARA"] = "0";
    }

    // Verificava espaco vazio no final do texto!

    if (e["DEPARA"] === "BDL_IP_FIXO") {
      e["DESIGNADOR"] += " " + e["DEPARA"];
    }

    if (e["TECNOLOGIA_ORIGEM"] === "FTTH") {
      e["DS_PRODUTO_TECNOLOGIA"] = "BANDA LARGA - FIBRA";
    } else if (e["TECNOLOGIA_ORIGEM"] === "VoIP FTTH") {
      e["DS_PRODUTO_TECNOLOGIA"] = "TERMINAL - FIBRA";
    } else if (
      e["TECNOLOGIA_ORIGEM"] === "FTTC" ||
      e["TECNOLOGIA_ORIGEM"] === "aDSL"
    ) {
      e["DS_PRODUTO_TECNOLOGIA"] = "BANDA LARGA - METALICO";
    } else if (
      e["TECNOLOGIA_ORIGEM"] === "COBRE" ||
      e["TECNOLOGIA_ORIGEM"] === "WLL"
    ) {
      e["DS_PRODUTO_TECNOLOGIA"] = "TERMINAL - METALICO";
    } else if (e["TECNOLOGIA_ORIGEM"] === "IPTV") {
      e["DS_PRODUTO_TECNOLOGIA"] = "TV";
    } else {
      e["DS_PRODUTO_TECNOLOGIA"] = "OUTROS";
    }

    e["VL_FAT_BRUTO"] = parseFloat(e["VL_FAT_BRUTO"]);
  });

  // Ordenar
  content.sort((b, a) => {
    return String(a["M"]).localeCompare(String(b["M"]));
  });

  content.sort((a, b) => {
    return String(a["DOCUMENTO"]).localeCompare(String(b["DOCUMENTO"]));
  });

  // Função para unificar os objetos e somar a chave 4 quando a chave 3 for igual
  const resultado = content.reduce((acc, obj) => {
    const chave = obj["DESIGNADOR"];
    if (!acc[chave]) {
      acc[chave] = { ...obj }; // Cria uma cópia do objeto
    } else {
      acc[chave]["VL_FAT_BRUTO"] += parseFloat(obj["VL_FAT_BRUTO"]); // Soma a chave 4
    }
    return acc;
  }, {});

  // Funções de comparação para ordenar
  const compare = (b, a) => {
    if (a["M"] < b["M"]) {
      return -1;
    }

    if (a["M"] > b["M"]) {
      return 1;
    }
    return 0;
  };

  const compare2 = (a, b) => {
    if (a["DOCUMENTO"] < b["DOCUMENTO"]) {
      return -1;
    }
    if (a["DOCUMENTO"] > b["DOCUMENTO"]) {
      return 1;
    }
    return 0;
  };

  // Ordenar
  const certo1 = Object.values(resultado).sort(compare);
  const certo2 = certo1.sort(compare2);

  console.log(certo2);

  // Mantêm apenas duas casas decimais
  certo2.forEach((e) => {
    e["VL_FAT_BRUTO"] = e["VL_FAT_BRUTO"].toFixed(2);
  });

  console.log("Terminou");

  // Suponha que processedData seja o array de objetos que você quer salvar
  const workbook = XLSX.utils.book_new(); // Cria um novo workbook
  const worksheet = XLSX.utils.json_to_sheet(certo2); // Converte o array de objetos para uma worksheet

  // Adiciona a worksheet ao workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, "Dados Processados");

  // Gera o arquivo XLSX em um formato adequado para download
  const xlsxBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });

  // Cria um blob a partir do buffer
  const blob = new Blob([xlsxBuffer], { type: "application/octet-stream" });

  // Cria um link de download e dispara o download automaticamente
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "meu_arquivo.xlsx"; // Nome do arquivo
  document.body.appendChild(link);
  link.click();
  loader.setAttribute("style", "display: none");
  pProcess.setAttribute("style", "display: none");
}
