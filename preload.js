const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("opsElectron", {
  convertXls: (payload) => ipcRenderer.invoke("convert-xls", payload),
  importXlsx: (payload) => ipcRenderer.invoke("import-xlsx", payload),
  clearTable: () => ipcRenderer.invoke("clear-table"),
  downloadJasperPdf: (payload) => ipcRenderer.invoke("download-jasper-pdf", payload),
});
