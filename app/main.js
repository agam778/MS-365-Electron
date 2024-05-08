import { app, Menu, BrowserWindow, dialog, nativeImage } from "electron";
import { clearActivity, setActivity, loginToRPC } from './rpc.js';
import { initialize, trackEvent } from "@aptabase/electron/main";
import { ElectronBlocker } from "@cliqz/adblocker-electron";
import { getValue } from './store.js';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { join } from "path";

import Windows from "./useragents.json" with { type: "json" };
import checkInternetConnected from "check-internet-connected";
import contextMenu from "electron-context-menu";
import updaterpkg from "electron-updater";
import ElectronDl from "electron-dl";
import menulayout from './menu.js'
import logpkg from 'electron-log';

const { transports, log: _log, functions } = logpkg;
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const { autoUpdater } = updaterpkg;

transports.file.level = "verbose";
console.log = _log;
Object.assign(console, functions);

initialize("A-US-2528580917");

function createWindow() {
  const enterpriseOrNormal = getValue("enterprise-or-normal");
  const partition = enterpriseOrNormal === "https://microsoft365.com/?auth=1" ? "persist:personal" : "persist:work";
  
  const win = new BrowserWindow({
    width: 1181,
    height: 670,
    icon: join(__dirname, "/assets/icons/png/1024x1024.png"),
    show: false,
    webPreferences: {
      nodeIntegration: true,
      devTools: true,
      partition: partition,
    },
  });  
  
  win.setAutoHideMenuBar(getValue("autohide-menubar") === "true");

  const splash = new BrowserWindow({
    width: 810,
    height: 610,
    transparent: true,
    frame: false,
    icon: join(__dirname, "/assets/icons/png/1024x1024.png"),
  });

  splash.loadURL(`https://agam778.github.io/MS-365-Electron/loading`);
  win.loadURL(
    `${getValue("enterprise-or-normal") || "https://microsoft365.com/?auth=1"}`,
    {
      userAgent: getValue("useragentstring") || Windows,
    }
  );

  win.webContents.on("did-finish-load", () => {
    splash.destroy();
    win.show();
    trackEvent("app_started");
    if (getValue("discordrpcstatus") === "true") {
      setActivity(`On "${win.webContents.getTitle()}"`);
    }
    if (getValue("blockadsandtrackers") === "true") {
      ElectronBlocker.fromPrebuiltAdsAndTracking(fetch).then((blocker) => {
        blocker.enableBlockingInSession(win.webContents.session);
      });
    }
  });
}

ElectronDl({
  dlPath: "./downloads",
  onStarted: (item) => {
    dialog.showMessageBox({
      type: "info",
      title: "Downloading File",
      message: `Downloading "${item.getFilename()}" to "${item.getSavePath()}"`,
      buttons: ["OK"],
    });
  },
  onCompleted: () => {
    dialog.showMessageBox({
      type: "info",
      title: "Download Completed",
      message: `Downloading Completed! Please check your "Downloads" folder.`,
      buttons: ["OK"],
    });
  },
  onError: (item) => {
    dialog.showMessageBox({
      type: "error",
      title: "Download failed",
      message: `Downloading "${item.getFilename()}" failed :(`,
      buttons: ["OK"],
    });
  },
});

contextMenu({
  showInspectElement: false,
  showServices: false,
});

Menu.setApplicationMenu(Menu.buildFromTemplate(menulayout));

app.on("ready", () => {
  createWindow();
});

app.on("web-contents-created", (event, contents) => {
  contents.setWindowOpenHandler(({ url }) => {
    if (getValue("websites-in-new-window") === "false") {
      if (url.includes("page=Download")) {
        return { action: "allow" };
      } else {
        BrowserWindow.getFocusedWindow()
          .loadURL(url)
          .catch((err) => {
            // do not show error
          });
        if (getValue("discordrpcstatus") === "true") {
          setActivity(
            `On "${BrowserWindow.getFocusedWindow().webContents.getTitle()}"`
          );
        }
        return { action: "deny" };
      }
    } else {
      if (getValue("discordrpcstatus") === "true") {
        setActivity(
          `On "${BrowserWindow.getFocusedWindow().webContents.getTitle()}"`
        );
      }
      return { action: "allow", overrideBrowserWindowOptions: { width: 1081, height: 610 } };
    }
  });
  contents.on("did-finish-load", () => {
    if (getValue("dynamicicons") === "true") {
      if (BrowserWindow.getFocusedWindow()) {
        if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("&ithint=file%2cpptx") ||
          BrowserWindow.getFocusedWindow()
            .webContents.getTitle()
            .includes(".pptx")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              join(__dirname, "../assets/icons/apps/powerpoint-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              join(__dirname, "../assets/icons/apps/powerpoint.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "PowerPoint");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("&ithint=file%2cdocx") ||
          BrowserWindow.getFocusedWindow()
            .webContents.getTitle()
            .includes(".docx")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              join(__dirname, "../assets/icons/apps/word-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              join(__dirname, "../assets/icons/apps/word.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "Word");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("&ithint=file%2cxlsx") ||
          BrowserWindow.getFocusedWindow()
            .webContents.getTitle()
            .includes(".xlsx")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              join(__dirname, "../assets/icons/apps/excel-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              join(__dirname, "../assets/icons/apps/excel.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "Excel");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("outlook.live.com") ||
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("outlook.office.com")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              join(__dirname, "../assets/icons/apps/outlook-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              join(__dirname, "../assets/icons/apps/outlook.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "Outlook");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("onedrive.live.com") ||
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("onedrive.aspx")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              join(__dirname, "../assets/icons/apps/onedrive-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              join(__dirname, "../assets/icons/apps/onedrive.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "OneDrive");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("teams.live.com")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              join(__dirname, "../assets/icons/apps/teams-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              join(__dirname, "../assets/icons/apps/teams.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "Teams");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("&ithint=onenote")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              join(__dirname, "../assets/icons/apps/onenote-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              join(__dirname, "../assets/icons/apps/onenote.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "OneNote");
            });
          }
        } else {
          if (process.platform === "darwin") {
            app.dock.setIcon(null);
          } else {
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(null, "");
            });
          }
        }
      }
    }
    contents.insertCSS(
      `
      ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
      }

      ::-webkit-scrollbar-track {
        background: transparent;
      }

      ::-webkit-scrollbar-thumb {
        background: transparent;
        border-radius: 5px;
      }

      ::-webkit-scrollbar-thumb:hover {
        background: #555;
      }      
      `
    );
  });
});

app.on("browser-window-created", (event, window) => {
  if (getValue("autohide-menubar") === "true") {
    window.setAutoHideMenuBar(true);
  } else {
    window.setAutoHideMenuBar(false);
  }
  window.webContents.on("did-finish-load", () => {
    if (getValue("discordrpcstatus") === "true") {
      setActivity(`On "${window.webContents.getTitle()}"`);
    }
  });
  if (getValue("blockadsandtrackers") === "true") {
    ElectronBlocker.fromPrebuiltAdsAndTracking(fetch).then((blocker) => {
      blocker.enableBlockingInSession(window.webContents.session);
    });
  }
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
  if (process.platform === "darwin") {
    app.dock.setIcon(null);
  }
  clearActivity();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

app.on("ready", function () {
  checkInternetConnected().catch(() => {
    const options = {
      type: "warning",
      buttons: ["Ok"],
      defaultId: 2,
      title: "Warning",
      message: "You appear to be offline!",
      detail:
        "Please check your Internet Connectivity. This app cannot run without an Internet Connection!",
    };
    dialog.showMessageBox(null, options, (response) => {
      console.log(response);
    });
  });
  if (getValue("autoupdater") === "true") {
    autoUpdater.checkForUpdatesAndNotify();
  }
  if (getValue("discordrpcstatus") === "true") {
    loginToRPC();
    setActivity(`Opening Microsoft 365...`);
  }
});
