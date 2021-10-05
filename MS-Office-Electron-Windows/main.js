const { app, Menu, BrowserWindow, globalShortcut } = require("electron");
const { dialog } = require("electron");
const openAboutWindow = require("about-window").default;
const join = require("path").join;
const isMac = process.platform === "darwin";
const path = require("path");
const electron = require("electron");
const { autoUpdater } = require("electron-updater");
const isOnline = require("is-online");

const template = [
  // { role: 'appMenu' }
  ...(isMac
    ? [
        {
          label: app.name,
          submenu: [
            { role: "about" },
            { type: "separator" },
            { role: "services" },
            { type: "separator" },
            { role: "hide" },
            { role: "hideothers" },
            { role: "unhide" },
            { type: "separator" },
            { role: "quit" },
          ],
        },
      ]
    : []),
  // { role: 'fileMenu' }
  {
    label: "Application",
    submenu: [
      {
        label: "About MS Office - Electron",
        click: () =>
          openAboutWindow({
            icon_path:
              "https://github.com/agam778/MS-Office-Electron/blob/main/icon2.png?raw=true",
            product_name: "MS Office - Electron",
            copyright: "Copyright (c) 2021 Agampreet Singh Bajaj",
            package_json_dir: __dirname,
            bug_report_url:
              "https://github.com/agam778/Microsoft-Office-Electron/issues/",
            bug_link_text: "Report an issue",
            adjust_window_size: "2",
            show_close_button: "Close",
          }),
      },
      {
        label: "Learn More",
        click: async () => {
          const { shell } = require("electron");
          await shell.openExternal(
            "https://github.com/agam778/MS-Office-Electron"
          );
        },
      },
      { type: "separator" },
      {
        role: "quit",
        accelerator: process.platform === "darwin" ? "Ctrl+Q" : "Ctrl+Q",
      },
    ],
  },
  // { role: 'editMenu' }
  {
    label: "Edit",
    submenu: [
      { role: "undo" },
      { role: "redo" },
      { type: "separator" },
      { role: "cut" },
      { role: "copy" },
      { role: "paste" },
      ...(isMac
        ? [
            { role: "pasteAndMatchStyle" },
            { role: "delete" },
            { role: "selectAll" },
            { type: "separator" },
            {
              label: "Speech",
              submenu: [{ role: "startSpeaking" }, { role: "stopSpeaking" }],
            },
          ]
        : [{ role: "delete" }, { type: "separator" }, { role: "selectAll" }]),
    ],
  },
  // { role: 'viewMenu' }
  {
    label: "View",
    submenu: [
      { role: "reload" },
      { role: "forceReload" },
      { type: "separator" },
      { role: "resetZoom" },
      {
        role: "zoomIn",
        accelerator: process.platform === "darwin" ? "Control+=" : "Control+=",
      },
      { role: "zoomOut" },
      { type: "separator" },
      { role: "togglefullscreen" },
    ],
  },
  // { role: 'windowMenu' }
  {
    label: "Window",
    submenu: [
      { role: "minimize" },
      { role: "zoom" },
      ...(isMac
        ? [
            { type: "separator" },
            { role: "front" },
            { type: "separator" },
            { role: "window" },
          ]
        : [{ role: "close" }]),
    ],
  },
];

const menu = Menu.buildFromTemplate(template);
Menu.setApplicationMenu(menu);

function createWindow() {
  const win = new BrowserWindow({
    width: 1181,
    height: 670,
    icon: "./icon.ico",
    webPreferences: {
      nodeIntegration: true,
      devTools: false,
      autoHideMenuBar: true,
    },
  });

  win.loadURL("https://agam778.gitlab.io/Microsoft-Office-Electron/", {
    userAgent:
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.71 Safari/537.36",
  });
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

try {
  require("electron-reloader")(module);
} catch (_) {}

app.on("ready", function () {
  isOnline().then((online) => {
    if (online) {
      console.log("You are connected to the internet!");
    } else {
      const options = {
        type: "warning",
        buttons: ["Ok"],
        defaultId: 2,
        title: "Warning",
        message: "You appear to be offline!",
        detail:
          "Please check your Internet Connectivity. This app cannot run without an Internet Connection!",
      };

      dialog.showMessageBox(null, options, (response, checkboxChecked) => {
        console.log(response);
      });
    }
  });
  autoUpdater.checkForUpdatesAndNotify();
});
