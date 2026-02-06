// src/js/fdvBasica_novo.js (AJUSTADO FINAL)

(function () {
  const loaderNew = document.getElementById("loaderNew");
  const pNew = document.getElementById("pNew");
  const btnExportarNew = document.getElementById("exportBtnNew");
  const fileLabelNew = document.getElementById("labelFileInputNew");
  const fileInputNew = document.getElementById("fileInputNew");

  // Se a página não tiver o novo, não faz nada
  if (!loaderNew || !pNew || !btnExportarNew || !fileLabelNew || !fileInputNew) return;

  // Colunas finais (ordem) - NOVO LAYOUT
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

  // Extras (opcional)
  const EXTRA_COLUMNS = ["STATUS_ATUAL", "DT_STATUS", "NOTA_ATUALIZACAO"];
  const HEADER_ORDER = [...BASE_COLUMNS, ...EXTRA_COLUMNS];

  let resumoData = [];

  function setLoading(isLoading) {
    loaderNew.style.display = isLoading ? "block" : "none";
    pNew.style.display = isLoading ? "block" : "none";
  }

  // garante estado inicial
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

  function readFileAsJson(file, cb, onErr) {
    const name = (file.name || "").toLowerCase();

    try {
      if (name.endsWith(".csv")) {
        const reader = new FileReader();
        reader.onerror = () => onErr && onErr(reader.error || new Error("Erro ao ler CSV"));
        reader.onload = function (event) {
          try {
            const csvText = String(event.target.result || "");
            const delim = detectCsvDelimiter(csvText);
            const wb = XLSX.read(csvText, { type: "string", FS: delim });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            cb(json);
          } catch (e) {
            onErr && onErr(e);
          }
        };
        reader.readAsText(file, "utf-8");
        return;
      }

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

  function clean(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/[\x00-\x1F\x7F]/g, "").trim();
  }

  // Normaliza cabeçalho: remove espaços, troca por underline, deixa em maiúsculo
  function normalizeKey(k) {
    return clean(k)
      .toUpperCase()
      .replace(/\s+/g, "_")
      .replace(/-+/g, "_");
  }

  // Converte cada linha para ter chaves normalizadas
  function normalizeRows(rows) {
    return rows.map((r) => {
      const out = {};
      Object.keys(r || {}).forEach((k) => {
        out[normalizeKey(k)] = r[k];
      });
      return out;
    });
  }

  // Regra: pega exatamente as colunas do novo layout; extras ficam vazias
  function buildResumoRows(rows) {
    const normalized = normalizeRows(rows);

    return normalized.map((r) => {
      const out = {};
      BASE_COLUMNS.forEach((c) => (out[c] = clean(r[c])));

      // Extras (opcional)
      out.STATUS_ATUAL = "";
      out.DT_STATUS = "";
      out.NOTA_ATUALIZACAO = "";

      return out;
    });
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
