const { app, Menu, BrowserWindow, dialog, nativeImage } = require("electron");
const { autoUpdater } = require("electron-updater");
const checkInternetConnected = require("check-internet-connected");
const ElectronDl = require("electron-dl");
const contextMenu = require("electron-context-menu");
const path = require("path");
const log = require("electron-log");
const { setActivity, loginToRPC, clearActivity } = require("./rpc");
const useragents = require("./useragents.json");
const { ElectronBlocker } = require("@cliqz/adblocker-electron");
const { getValue } = require("./store");
const { menulayout } = require("./menu");

log.transports.file.level = "verbose";
console.log = log.log;
Object.assign(console, log.functions);

function createWindow() {
  const win = new BrowserWindow({
    width: 1181,
    height: 670,
    icon: path.join(__dirname, "/assets/icons/png/1024x1024.png"),
    show: false,
    webPreferences: {
      nodeIntegration: true,
      devTools: true,
    },
  });

  if (getValue("autohide-menubar") === "true") {
    win.setAutoHideMenuBar(true);
  } else {
    win.setAutoHideMenuBar(false);
  }

  const splash = new BrowserWindow({
    width: 810,
    height: 610,
    transparent: true,
    frame: false,
    alwaysOnTop: true,
    icon: path.join(__dirname, "/assets/icons/png/1024x1024.png"),
  });

  splash.loadURL(`https://agam778.github.io/MS-365-Electron/loading`);
  win.loadURL(
    `${getValue("enterprise-or-normal") || "https://microsoft365.com/?auth=1"}`,
    {
      userAgent: getValue("useragentstring") || useragents.Windows,
    }
  );

  win.webContents.on("did-finish-load", () => {
    splash.destroy();
    win.show();
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
      return { action: "allow" };
    }
  });
  contents.on("did-finish-load", () => {
    if (getValue("dynamicicons") === "true") {
      if (BrowserWindow.getFocusedWindow()) {
        if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("&ithint=file%2cpptx")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              path.join(__dirname, "../assets/icons/apps/powerpoint-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              path.join(__dirname, "../assets/icons/apps/powerpoint.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "PowerPoint");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("&ithint=file%2cdocx")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              path.join(__dirname, "../assets/icons/apps/word-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              path.join(__dirname, "../assets/icons/apps/word.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "Word");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("&ithint=file%2cxlsx")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              path.join(__dirname, "../assets/icons/apps/excel-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              path.join(__dirname, "../assets/icons/apps/excel.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "Excel");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("outlook.live.com")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              path.join(__dirname, "../assets/icons/apps/outlook-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              path.join(__dirname, "../assets/icons/apps/outlook.png")
            );
            BrowserWindow.getAllWindows().forEach((window) => {
              window.setOverlayIcon(nimage, "Outlook");
            });
          }
        } else if (
          BrowserWindow.getFocusedWindow()
            .webContents.getURL()
            .includes("onedrive.live.com")
        ) {
          if (process.platform === "darwin") {
            app.dock.setIcon(
              path.join(__dirname, "../assets/icons/apps/onedrive-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              path.join(__dirname, "../assets/icons/apps/onedrive.png")
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
              path.join(__dirname, "../assets/icons/apps/teams-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              path.join(__dirname, "../assets/icons/apps/teams.png")
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
              path.join(__dirname, "../assets/icons/apps/onenote-mac.png")
            );
          } else if (process.platform === "win32") {
            let nimage = nativeImage.createFromPath(
              path.join(__dirname, "../assets/icons/apps/onenote.png")
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
  });
});

app.on("browser-window-created", (event, window) => {
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
