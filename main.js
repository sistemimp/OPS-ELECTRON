const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const { autoUpdater } = require("electron-updater");
const path = require("path");
const fs = require("fs");
const https = require("https");
const XLSX = require("xlsx");
const mysql = require("mysql2/promise");

const DB_CONFIG = {
  host: "82.223.30.31",
  port: 3306,
  user: "mediaprint",
  password: "M3d1aPr1ntDB@",
  database: "OpsMp",
};

const JASPER_AUTH = {
  user: "jasperadmin",
  password: "jasperadmin",
};

const JASPER_REPORTS = new Map([
  [
    "OPS_AR1",
    "https://jaspersoft.mediaprint.it/jasperserver/rest_v2/reports/Mediaprint/Clienti/OPS/Layout/OPS_AR1.pdf",
  ],
  [
    "OPS_CartolineAR",
    "https://jaspersoft.mediaprint.it/jasperserver/rest_v2/reports/Mediaprint/Clienti/OPS/Layout/OPS_CartolineAR.pdf",
  ],
]);

function createWindow() {
  const win = new BrowserWindow({
    width: 900,
    height: 720,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.loadFile("index.html");
}

function setupAutoUpdate() {
  if (!app.isPackaged) {
    return;
  }

  const feedUrl = process.env.OPS_AUTOUPDATE_URL;
  if (feedUrl) {
    autoUpdater.setFeedURL({ provider: "generic", url: feedUrl });
  }

  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on("error", (error) => {
    console.error("Auto update error:", error);
  });

  autoUpdater.on("update-downloaded", () => {
    autoUpdater.quitAndInstall();
  });

  autoUpdater.checkForUpdates().catch((error) => {
    console.error("Auto update check failed:", error);
  });
}

app.whenReady().then(() => {
  createWindow();
  setupAutoUpdate();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

function getBasicAuthHeader() {
  const token = Buffer.from(`${JASPER_AUTH.user}:${JASPER_AUTH.password}`).toString("base64");
  return `Basic ${token}`;
}

function downloadPdfWithRedirect(url, destPath, maxRedirects = 5) {
  return new Promise((resolve, reject) => {
    if (maxRedirects < 0) {
      reject(new Error("Troppi redirect durante il download."));
      return;
    }

    const requestUrl = new URL(url);
    const options = {
      headers: {
        Authorization: getBasicAuthHeader(),
      },
    };

    https
      .get(requestUrl, options, (res) => {
        const { statusCode, headers } = res;
        if (statusCode && statusCode >= 300 && statusCode < 400 && headers.location) {
          res.resume();
          downloadPdfWithRedirect(headers.location, destPath, maxRedirects - 1)
            .then(resolve)
            .catch(reject);
          return;
        }

        if (statusCode !== 200) {
          res.resume();
          reject(new Error(`Download fallito (HTTP ${statusCode}).`));
          return;
        }

        const fileStream = fs.createWriteStream(destPath);
        res.pipe(fileStream);

        fileStream.on("finish", () => {
          fileStream.close(resolve);
        });

        fileStream.on("error", (err) => {
          fs.unlink(destPath, () => reject(err));
        });
      })
      .on("error", reject);
  });
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatDateToText(date) {
  if (!date) return "";
  if (date instanceof Date && !isNaN(date.getTime())) {
    return `${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}/${date.getFullYear()}`;
  }
  return "";
}

function excelSerialToDate(serial) {
  const parsed = XLSX.SSF.parse_date_code(serial);
  if (!parsed) return null;
  return new Date(parsed.y, parsed.m - 1, parsed.d);
}

function parseDateInput(value) {
  if (!value || typeof value !== "string") return null;
  const parts = value.split("-");
  if (parts.length !== 3) return null;
  const [y, m, d] = parts.map((p) => Number(p));
  if (!y || !m || !d) return null;
  return new Date(y, m - 1, d);
}

function normalizeCell(value, headerLower) {
  if (value === null || value === undefined) return "";

  if (headerLower.includes("data")) {
    if (value instanceof Date) return formatDateToText(value);
    if (typeof value === "number") {
      const d = excelSerialToDate(value);
      return d ? formatDateToText(d) : String(value);
    }
    if (typeof value === "string") {
      // Best effort: keep as-is if already text
      return value;
    }
    return String(value);
  }

  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (value instanceof Date) return formatDateToText(value);
  return String(value);
}

function computeBarcode23i(code) {
  const raw = code === null || code === undefined ? "" : String(code).trim();
  const codeLoc = raw.padStart(11, "0").slice(-11);
  let sommaPari = 0;
  let sommaDispari = 0;

  for (let i = 1; i <= 11; i += 1) {
    const digit = Number(codeLoc[i - 1] || "0");
    if (i % 2 === 0) sommaPari += digit;
    else sommaDispari += digit;
  }

  let sommaResto = sommaPari * 11 + sommaDispari;
  let somma = Math.floor(sommaResto / 100);
  sommaResto = sommaResto % 100;
  somma += Math.floor(sommaResto / 10);
  sommaResto = sommaResto % 10;
  somma += sommaResto;
  const checkDigit = somma % 10;

  return `${codeLoc}-${checkDigit}`;
}

function mapVerificatore(value) {
  if (value === null || value === undefined) return "";
  const raw = String(value).trim();
  if (!raw) return "";
  switch (raw) {
    case "Enzo Masciarelli":
      return "Sig. Enzo Masciarelli - 348/8274706";
    case "Gennaro Falcone":
      return "Sig. Gennaro Falcone - 348/8274708";
    case "Eliano Manzone":
      return "Sig. Eliano Manzone - 349/7604977";
    case "Alessandro Fanci":
      return "Dott. Alessandro Fanci - 348/6722097";
    case "Celestino Visconti":
      return "Sig. Celestino Visconti - 348/6621987";
    case "Stefano Battista":
      return "Dott. Stefano Battista - 348/8274705";
    default:
      return raw;
  }
}

function sanitizeHeader(name) {
  const base = String(name || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9_]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
  return base || "col";
}

function makeUniqueHeaders(headers) {
  const seen = new Map();
  return headers.map((h) => {
    const key = h.toLowerCase();
    const count = seen.get(key) || 0;
    seen.set(key, count + 1);
    if (count === 0) return h;
    return `${h}_${count + 1}`;
  });
}

ipcMain.handle("convert-xls", async (_event, payload) => {
  const { protocolNumber, protocolDate, lavorazioneCode, barcodeStart, inputPath } = payload;

  if (!inputPath) throw new Error("Seleziona un file .xls");

  const protocolDateText = formatDateToText(parseDateInput(protocolDate));

  const data = fs.readFileSync(inputPath);
  const wb = XLSX.read(data, { type: "buffer", cellDates: true });
  const firstSheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (rows.length === 0) throw new Error("Il file non contiene righe.");

  const headers = rows[0].map((h) => (h === null || h === undefined ? "" : String(h)));
  const headersLower = headers.map((h) => h.toLowerCase());

  const extraHeaders = [
    "progressivo",
    "CODICE_AR",
    "CODICE_AR23i",
    "Barcode23i",
    "verificatore ok",
    "n.protocollo",
    "data protocollo",
    "codice lavorazione",
    "omologazione",
  ];

  const output = [];
  output.push([...headers, ...extraHeaders]);

  const barcodeBase = Number.parseInt(barcodeStart, 10);

  const verificatoreIdx = headersLower.indexOf("verificatore");

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const normalized = headers.map((_, idx) => normalizeCell(row[idx], headersLower[idx] || ""));

    const progressivo = i;
    const barcode = Number.isFinite(barcodeBase) ? barcodeBase + (i - 1) : "";
    const codiceAr23i = computeBarcode23i(barcode);
    const barcode23i = codiceAr23i.replace("-", "");
    const verificatoreValue =
      verificatoreIdx >= 0 ? mapVerificatore(row[verificatoreIdx]) : "";
    const extra = [
      progressivo,
      barcode,
      codiceAr23i,
      barcode23i,
      verificatoreValue,
      protocolNumber || "",
      protocolDateText,
      lavorazioneCode || "",
      "DCOCC0015",
    ];

    output.push([...normalized, ...extra]);
  }

  const outWb = XLSX.utils.book_new();
  const outSheet = XLSX.utils.aoa_to_sheet(output);

  // Force all cells to string to keep text representation
  const range = XLSX.utils.decode_range(outSheet["!ref"]);
  const progressivoCol = headers.length;
  const codiceArCol = headers.length + 1;

  for (let R = range.s.r; R <= range.e.r; R += 1) {
    for (let C = range.s.c; C <= range.e.c; C += 1) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = outSheet[cellAddress];
      if (!cell) continue;
      const isHeaderRow = R === 0;
      if (!isHeaderRow && (C === progressivoCol || C === codiceArCol)) {
        const n = Number(cell.v);
        cell.t = "n";
        cell.v = Number.isFinite(n) ? n : 0;
        continue;
      }
      cell.t = "s";
      cell.v = cell.v === null || cell.v === undefined ? "" : String(cell.v);
    }
  }

  XLSX.utils.book_append_sheet(outWb, outSheet, "Output");

  const defaultName = path.basename(inputPath, path.extname(inputPath)) + "_convertito.xlsx";
  const filePath = path.join(path.dirname(inputPath), defaultName);

  XLSX.writeFile(outWb, filePath);
  return { saved: true, filePath };
});

ipcMain.handle("import-xlsx", async (_event, payload) => {
  const { inputPath } = payload;
  if (!inputPath) throw new Error("Seleziona un file .xlsx.");

  const data = fs.readFileSync(inputPath);
  const wb = XLSX.read(data, { type: "buffer", cellDates: false });
  const firstSheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (rows.length < 2) throw new Error("Il file non contiene dati da importare.");

  const rawHeaders = rows[0].map((h) => (h === null || h === undefined ? "" : String(h)));
  const sanitized = rawHeaders.map(sanitizeHeader);
  const headers = makeUniqueHeaders(sanitized);
  const dataRows = rows.slice(1).filter((r) => r && r.length > 0);

  const columns = headers.map((h) => `\`${h.replace(/`/g, "``")}\``).join(", ");
  const placeholders = headers.map(() => "?").join(", ");
  const sql = `INSERT INTO \`OPS-AR\` (${columns}) VALUES (${placeholders})`;

  const connection = await mysql.createConnection(DB_CONFIG);
  try {
    let inserted = 0;
    for (const row of dataRows) {
      const values = headers.map((_, idx) => {
        const v = row[idx];
        if (v === null || v === undefined) return "";
        return String(v);
      });
      await connection.execute(sql, values);
      inserted += 1;
    }
    return { inserted };
  } finally {
    await connection.end();
  }
});

ipcMain.handle("clear-table", async () => {
  const connection = await mysql.createConnection(DB_CONFIG);
  try {
    const [result] = await connection.execute("DELETE FROM `OPS-AR`");
    const deleted = result.affectedRows || 0;
    return { deleted };
  } finally {
    await connection.end();
  }
});

let lastJasperDir = "";

ipcMain.handle("download-jasper-pdf", async (_event, payload) => {
  const { reportKey, sourcePath } = payload || {};
  const url = JASPER_REPORTS.get(reportKey);
  if (!url) throw new Error("Report Jasper non riconosciuto.");

  if (!sourcePath) {
    throw new Error("File XLS non trovato per determinare la cartella.");
  }

  const now = new Date();
  const timestamp = [
    now.getFullYear(),
    pad2(now.getMonth() + 1),
    pad2(now.getDate()),
    "_",
    pad2(now.getHours()),
    pad2(now.getMinutes()),
    pad2(now.getSeconds()),
  ].join("");

  const defaultName = `${reportKey}_${timestamp}.pdf`;
  const folder = path.dirname(sourcePath);
  lastJasperDir = folder;
  const destPath = path.join(folder, defaultName);

  await downloadPdfWithRedirect(url, destPath);
  return { saved: true, filePath: destPath };
});
