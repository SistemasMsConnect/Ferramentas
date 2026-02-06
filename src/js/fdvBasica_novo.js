// src/js/fdvBasica_novo.js (REVISTO - mantém acentos)

(function () {
  const loaderNew = document.getElementById("loaderNew");
  const pNew = document.getElementById("pNew");
  const btnExportarNew = document.getElementById("exportBtnNew");
  const fileLabelNew = document.getElementById("labelFileInputNew");
  const fileInputNew = document.getElementById("fileInputNew");

  if (!loaderNew || !pNew || !btnExportarNew || !fileLabelNew || !fileInputNew) return;

  const BASE_COLUMNS = [
    "NR_DOCUMENTO",
    "ID_SIMULACAO",
    "STATUS_SIMULACAO",
    "FILA_SIMPLIFIQUE",
    "CRITICA_SIMULACAO",
    "MOTIVO_NAOTRATAVEL_SIMULACAO",
    "FILA_ESTEIRA_SIEBEL",
    "FILA_ESTEIRA_WFM",
    "MOTIVO_PENDENCIA_WFM",
    "SUB_MOTIVO_WFM",
    "STATUS_ATIVIDADE_ZEUS",
    "NOTA_TECNICO",
    "NOT_DONE_ZEUS",
    "DATA_AGENDAMENTO_ZEUS",
    "INICIO_ATIVIDADE_ZEUS",
    "FIM_DA_ATIVIDADE_ZEUS",
    "NR_RPON",
    "NR_PON_VENDA",
    "NR_ORDEM",
    "DT_VENDA",
    "DT_ATIVACAO",
    "DT_CANCELAMENTO",
    "DT_DESCONEXAO",
    "SIT_PORTABILIDADE",
    "DT_INI_PORTABILIDADE",
    "DT_EXEC_PORT",
    "STATUS_BACKLOG",
    "DETALHAMENTO_BACKLOG",
    "TP_CANCELAMENTO",
    "MOTIVO_CANCELAMENTO",
    "BASE_ORIGEM",
    "DT_CARGA",
    "DT_CARGA_WFM_ZEUS",
  ];

  const EXTRA_COLUMNS = ["STATUS_ATUAL", "DT_STATUS", "NOTA_ATUALIZACAO"];
  const HEADER_ORDER = [...BASE_COLUMNS, ...EXTRA_COLUMNS];

  let resumoData = [];

  function setLoading(isLoading) {
    loaderNew.style.display = isLoading ? "block" : "none";
    pNew.style.display = isLoading ? "block" : "none";
  }
  setLoading(false);

  function downloadWorkbook(workbook, filename) {
    const wbout = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
      compression: true,
    });

    const blob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function detectCsvDelimiter(text) {
    const sample = text.split(/\r?\n/).slice(0, 5).join("\n");
    const commaCount = (sample.match(/,/g) || []).length;
    const semiCount = (sample.match(/;/g) || []).length;
    return semiCount > commaCount ? ";" : ",";
  }

  function clean(v) {
    if (v === null || v === undefined) return "";
    // NÃO remove acento; só remove caracteres de controle e trim
    return String(v).replace(/[\x00-\x1F\x7F]/g, "").trim();
  }

  function normalizeKey(k) {
    return clean(k)
      .toUpperCase()
      .replace(/\s+/g, "_")
      .replace(/-+/g, "_");
  }

  function normalizeRows(rows) {
    return rows.map((r) => {
      const out = {};
      Object.keys(r || {}).forEach((k) => {
        out[normalizeKey(k)] = r[k];
      });
      return out;
    });
  }

  function buildResumoRows(rows) {
    const normalized = normalizeRows(rows);

    return normalized.map((r) => {
      const out = {};

      BASE_COLUMNS.forEach((c) => (out[c] = clean(r[c])));

      out.STATUS_ATUAL = "";
      out.DT_STATUS = "";
      out.NOTA_ATUALIZACAO = "";

      return out;
    });
  }

  // Heurística simples: se o texto tem muitos "�" (replacement char),
  // provavelmente decodificou errado.
  function looksBroken(text) {
    const rep = (text.match(/\uFFFD/g) || []).length; // "�"
    return rep > 0;
  }

  // Lê CSV com fallback de encoding para manter acentos
  function readCsvWithBestEncoding(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);

    // BOM UTF-8
    const hasUtf8Bom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;

    // tenta UTF-8 primeiro
    let textUtf8 = new TextDecoder("utf-8", { fatal: false }).decode(bytes);
    if (hasUtf8Bom) textUtf8 = textUtf8.replace(/^\uFEFF/, "");

    // se não aparenta quebrado, usa UTF-8
    if (!looksBroken(textUtf8)) return textUtf8;

    // fallback: Windows-1252 (muito comum no Brasil / Excel)
    try {
      const text1252 = new TextDecoder("windows-1252", { fatal: false }).decode(bytes);
      if (!looksBroken(text1252)) return text1252;
      return text1252; // mesmo se tiver algum "�", ainda tende a preservar mais
    } catch (e) {
      // fallback final: latin1
      return new TextDecoder("iso-8859-1", { fatal: false }).decode(bytes);
    }
  }

  function readFileAsJson(file, cb, onErr) {
    const name = (file.name || "").toLowerCase();

    try {
      if (name.endsWith(".csv")) {
        const reader = new FileReader();
        reader.onerror = () => onErr && onErr(reader.error || new Error("Erro ao ler CSV"));

        reader.onload = function (event) {
          try {
            const buffer = event.target.result; // ArrayBuffer
            const csvText = readCsvWithBestEncoding(buffer);

            const delim = detectCsvDelimiter(csvText);
            const wb = XLSX.read(csvText, { type: "string", FS: delim });

            const sheet = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            cb(json);
          } catch (e) {
            onErr && onErr(e);
          }
        };

        // Lê como ArrayBuffer (não como texto), para decidir encoding corretamente
        reader.readAsArrayBuffer(file);
        return;
      }

      // XLSX normal (já preserva acentos)
      const reader = new FileReader();
      reader.onerror = () => onErr && onErr(reader.error || new Error("Erro ao ler XLSX"));

      reader.onload = function (event) {
        try {
          const data = new Uint8Array(event.target.result);
          const wb = XLSX.read(data, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          cb(json);
        } catch (e) {
          onErr && onErr(e);
        }
      };

      reader.readAsArrayBuffer(file);
    } catch (e) {
      onErr && onErr(e);
    }
  }

  fileInputNew.addEventListener("change", function (event) {
    const file = event.target.files[0];
    if (!file) return;

    fileLabelNew.textContent = file.name;
    setLoading(true);

    readFileAsJson(
      file,
      (json) => {
        try {
          if (!json || !json.length) {
            alert("NOVO: arquivo vazio ou não foi possível ler.");
            resumoData = [];
            setLoading(false);
            return;
          }

          resumoData = buildResumoRows(json);

          console.log("[NOVO] origem:", json.length, "resumo:", resumoData.length);
          setLoading(false);
        } catch (err) {
          console.error("[NOVO] erro:", err);
          alert("Erro no NOVO. Veja o console.");
          resumoData = [];
          setLoading(false);
        }
      },
      (err) => {
        console.error("[NOVO] erro leitura:", err);
        alert("Erro ao ler o arquivo (NOVO). Veja o console.");
        resumoData = [];
        setLoading(false);
      }
    );
  });

  btnExportarNew.addEventListener("click", function () {
    if (!resumoData.length) {
      alert("NOVO: Nenhum dado para exportar. Selecione o arquivo primeiro.");
      return;
    }

    const ws = XLSX.utils.json_to_sheet(resumoData, { header: HEADER_ORDER });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Resumo");

    downloadWorkbook(wb, "RESULTADO_FIXABASICA.xlsx");
  });
})();
