const { app, Menu, BrowserWindow, dialog } = require("electron");
const { autoUpdater } = require("electron-updater");
const checkInternetConnected = require("check-internet-connected");
const ElectronDl = require("electron-dl");
const contextMenu = require("electron-context-menu");
const path = require("path");
const store = require("./store");
const log = require("electron-log");
const { setActivity, loginToRPC, clearActivity } = require("./rpc");
const useragents = require("./useragents.json");

log.transports.file.level = "verbose";
console.log = log.log;
Object.assign(console, log.functions);

ElectronDl({
  dlPath: "./downloads",
  onStarted: (item) => {
    BrowserWindow.getAllWindows().forEach((window) => {
      window.setOverlayIcon(
        __dirname + "/assets/icons/download.png",
        "Downloading"
      );
    });
    dialog.showMessageBox({
      type: "info",
      title: "Downloading File",
      message: `Downloading "${item.getFilename()}" to "${item.getSavePath()}"`,
      buttons: ["OK"],
    });
  },
  onCompleted: () => {
    BrowserWindow.getAllWindows().forEach((window) => {
      window.setOverlayIcon(
        __dirname + "/assets/icons/download-success.png",
        "Download Successful"
      );
    });
    dialog.showMessageBox({
      type: "info",
      title: "Download Completed",
      message: `Downloading Completed! Please check your "Downloads" folder.`,
      buttons: ["OK"],
    });
    setTimeout(() => {
      BrowserWindow.getAllWindows().forEach((window) => {
        window.setOverlayIcon(null, "");
      });
    }, 7000);
  },
  onError: (item) => {
    dialog.showMessageBox({
      type: "error",
      title: "Download failed",
      message: `Downloading "${item.getFilename()}" failed :(`,
      buttons: ["OK"],
    });
    BrowserWindow.getAllWindows().forEach((window) => {
      window.setOverlayIcon(
        __dirname + "/download-fail.png",
        "Download Failed"
      );
    });
    setTimeout(() => {
      BrowserWindow.getAllWindows().forEach((window) => {
        window.setOverlayIcon(null, "");
      });
    }, 7000);
  },
});

contextMenu({
  showInspectElement: false,
  showServices: false,
});

const { menulayout } = require("./menu");

const menu = Menu.buildFromTemplate(menulayout);
Menu.setApplicationMenu(menu);

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

  if (store.get("autohide-menubar") === "true") {
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
    `${
      store.get("enterprise-or-normal") || "https://microsoft365.com/?auth=1"
    }`,
    {
      userAgent: store.get("useragentstring") || useragents.Windows,
    }
  );

  require("./rpc.js");

  win.webContents.on("did-finish-load", () => {
    splash.destroy();
    win.show();
    if (store.get("discordrpcstatus") === "true") {
      setActivity(`On "${win.webContents.getTitle()}"`);
    }
  });
}

app.on("ready", () => {
  createWindow();
});

app.on("web-contents-created", (event, contents) => {
  contents.setWindowOpenHandler(({ url }) => {
    if (store.get("websites-in-new-window") === "false") {
      if (url.includes("page=Download")) {
        return { action: "allow" };
      } else {
        BrowserWindow.getFocusedWindow().loadURL(url);
        if (store.get("discordrpcstatus") === "true") {
          setActivity(
            `On "${BrowserWindow.getFocusedWindow().webContents.getTitle()}"`
          );
        }
        return { action: "deny" };
      }
    } else {
      if (store.get("discordrpcstatus") === "true") {
        setActivity(
          `On "${BrowserWindow.getFocusedWindow().webContents.getTitle()}"`
        );
      }
      return { action: "allow" };
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
  clearActivity();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

app.on("ready", function () {
  checkInternetConnected()
    .then(() => {
      console.log("You are connected to the internet!");
    })
    .catch(() => {
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
  autoUpdater.checkForUpdatesAndNotify();
  if (store.get("discordrpcstatus") === "true") {
    loginToRPC();
    setActivity(`Opening Microsoft 365...`);
  }
});
