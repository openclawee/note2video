const { app, BrowserWindow, dialog, ipcMain, shell } = require("electron");
const path = require("path");
const fs = require("fs/promises");

/** @type {BrowserWindow | null} */
let mainWindow = null;

function isDev() {
  return !app.isPackaged;
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 640,
    backgroundColor: "#0f1117",
    titleBarStyle: process.platform === "darwin" ? "hiddenInset" : "default",
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  if (isDev()) {
    mainWindow.loadURL("http://127.0.0.1:5173");
    mainWindow.webContents.openDevTools({ mode: "detach" });
  } else {
    const indexHtml = path.join(__dirname, "..", "dist", "index.html");
    mainWindow.loadFile(indexHtml);
  }

  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

ipcMain.handle("dialog:openProjectDir", async () => {
  const res = await dialog.showOpenDialog({
    properties: ["openDirectory", "createDirectory"],
    title: "选择 Note2Video 输出目录（含 manifest.json）",
  });
  if (res.canceled || !res.filePaths[0]) {
    return null;
  }
  return res.filePaths[0];
});

ipcMain.handle("fs:readText", async (_evt, filePath) => {
  return await fs.readFile(filePath, "utf-8");
});

ipcMain.handle("fs:writeText", async (_evt, filePath, text) => {
  await fs.writeFile(filePath, text, "utf-8");
});

ipcMain.handle("shell:openPath", async (_evt, targetPath) => {
  const err = await shell.openPath(targetPath);
  return err || null;
});

ipcMain.handle("path:join", async (_evt, root, ...parts) => {
  return path.join(root, ...parts);
});

ipcMain.handle("path:toFileUrl", async (_evt, filePath) => {
  let p = path.resolve(filePath);
  if (process.platform === "win32") {
    p = p.replace(/\\/g, "/");
    if (!p.startsWith("/")) {
      p = `/${p}`;
    }
    return `file://${p}`;
  }
  return `file://${encodeURI(p)}`;
});
