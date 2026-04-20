const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("n2v", {
  openProjectDir: () => ipcRenderer.invoke("dialog:openProjectDir"),
  readText: (filePath) => ipcRenderer.invoke("fs:readText", filePath),
  writeText: (filePath, text) => ipcRenderer.invoke("fs:writeText", filePath, text),
  openPath: (targetPath) => ipcRenderer.invoke("shell:openPath", targetPath),
  joinPath: (root, ...parts) => ipcRenderer.invoke("path:join", root, ...parts),
  toFileUrl: (filePath) => ipcRenderer.invoke("path:toFileUrl", filePath),
});
