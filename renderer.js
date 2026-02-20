const form = document.getElementById("convert-form");
const fileInput = document.getElementById("xls-file");
const statusEl = document.getElementById("status");
const importBtn = document.getElementById("import-btn");
const clearBtn = document.getElementById("clear-btn");
const convertBtn = document.getElementById("convert-btn");
const jasperAr1Btn = document.getElementById("jasper-ar1-btn");
const jasperCartolineBtn = document.getElementById("jasper-cartoline-btn");
let lastConvertedPath = "";
let lastSourceXlsPath = "";
const originalButtonHtml = new Map();

function setButtonLoading(btn, isLoading) {
  if (!btn) return;
  if (isLoading) {
    if (!originalButtonHtml.has(btn)) {
      originalButtonHtml.set(btn, btn.innerHTML);
    }
    btn.innerHTML = '<span class="spinner-border spinner-border-sm" aria-hidden="true"></span>';
    btn.disabled = true;
  } else {
    if (originalButtonHtml.has(btn)) {
      btn.innerHTML = originalButtonHtml.get(btn);
    }
    btn.disabled = false;
  }
}

function setStatus(text, type = "info") {
  statusEl.textContent = text;
  const typeMap = {
    info: "info",
    success: "success",
    warning: "warning",
    error: "danger",
  };
  const alertType = typeMap[type] || "info";
  statusEl.className = `alert alert-${alertType} mt-4 mb-0`;
}

function getFormValues() {
  const protocolNumber = document.getElementById("protocol-number").value.trim();
  const protocolDate = document.getElementById("protocol-date").value;
  const lavorazioneCode = document.getElementById("lavorazione-code").value.trim();
  const barcodeStart = document.getElementById("barcode-start").value.trim();

  return { protocolNumber, protocolDate, lavorazioneCode, barcodeStart };
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  setStatus("Conversione in corso...", "info");

  try {
    if (!fileInput) {
      setStatus("Input file XLS non trovato nella pagina.", "error");
      return;
    }
    const file = fileInput.files[0];
    if (!file) {
      setStatus("Seleziona un file .xls prima di continuare.", "error");
      return;
    }
    lastSourceXlsPath = file.path || lastSourceXlsPath;

    const { protocolNumber, protocolDate, lavorazioneCode, barcodeStart } = getFormValues();

    if (!protocolNumber || !protocolDate || !lavorazioneCode || !barcodeStart) {
      setStatus("Compila tutti i campi del form.", "error");
      return;
    }

    setButtonLoading(convertBtn, true);
    const result = await window.opsElectron.convertXls({
      protocolNumber,
      protocolDate,
      lavorazioneCode,
      inputPath: file.path,
      barcodeStart,
    });

    if (!result.saved) {
      setStatus("Operazione annullata.", "warning");
      return;
    }

    lastConvertedPath = result.filePath || "";
    setStatus(`File salvato: ${result.filePath}`, "success");
  } catch (err) {
    setStatus(err.message || "Errore durante la conversione.", "error");
  } finally {
    setButtonLoading(convertBtn, false);
  }
});

importBtn.addEventListener("click", async () => {
  setStatus("Import in corso...", "info");

  try {
    setButtonLoading(importBtn, true);
    const inputPath = lastConvertedPath;
    if (!inputPath) {
      setStatus("Converti prima un file.", "error");
      return;
    }

    const result = await window.opsElectron.importXlsx({ inputPath });

    setStatus(
      `Import completato. Righe inserite: ${result.inserted}`,
      "success"
    );
  } catch (err) {
    setStatus(err.message || "Errore durante l'import.", "error");
  } finally {
    setButtonLoading(importBtn, false);
  }
});

clearBtn.addEventListener("click", async () => {
  setStatus("Svuotamento in corso...", "info");

  try {
    setButtonLoading(clearBtn, true);
    const result = await window.opsElectron.clearTable();
    setStatus(`Tabella svuotata. Righe eliminate: ${result.deleted}`, "success");
  } catch (err) {
    setStatus(err.message || "Errore durante lo svuotamento.", "error");
  } finally {
    setButtonLoading(clearBtn, false);
  }
});

async function handleJasperDownload(reportKey, label, btn) {
  setStatus(`Download ${label} in corso...`, "info");

  try {
    setButtonLoading(btn, true);
    const inputPath = lastSourceXlsPath || fileInput?.files?.[0]?.path;
    if (!inputPath) {
      setStatus("Seleziona prima il file .xls per impostare la cartella.", "error");
      return;
    }

    const result = await window.opsElectron.downloadJasperPdf({
      reportKey,
      sourcePath: inputPath,
    });
    if (result.canceled) {
      setStatus("Operazione annullata.", "warning");
      return;
    }
    setStatus(`PDF salvato: ${result.filePath}`, "success");
  } catch (err) {
    setStatus(err.message || "Errore durante il download.", "error");
  } finally {
    setButtonLoading(btn, false);
  }
}

jasperAr1Btn.addEventListener("click", () => {
  handleJasperDownload("OPS_AR1", "OPS_AR1.pdf", jasperAr1Btn);
});

jasperCartolineBtn.addEventListener("click", () => {
  handleJasperDownload("OPS_CartolineAR", "OPS_CartolineAR.pdf", jasperCartolineBtn);
});
