const { app, Menu, BrowserWindow, dialog, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const openAboutWindow = require("about-window").default;
const checkInternetConnected = require("check-internet-connected");
const axios = require("axios");
const ElectronDl = require("electron-dl");
const contextMenu = require("electron-context-menu");
const path = require("path");
const Store = require("electron-store");
const log = require("electron-log");

const store = new Store();

const RPC = require("discord-rpc");
const clientId = "942637872530460742";
const rpc = new RPC.Client({ transport: "ipc" });

const windowsuseragent =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36";
const macuseragent =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_2_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36";
const linuxuseragent =
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36";

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

const menulayout = [
  {
    label: "Application",
    submenu: [
      {
        label: "About MS-365-Electron",
        click: () =>
          openAboutWindow({
            icon_path:
              "https://raw.githubusercontent.com/agam778/MS-365-Electron/main/assets/about.png",
            product_name: "MS-365-Electron",
            copyright:
              "Copyright (c) 2021-2023 Agampreet Singh\nOffice, the name, website, images/icons\nare the intellectual properties of Microsoft.",
            package_json_dir: __dirname,
            bug_report_url:
              "https://github.com/agam778/MS-365-Electron/issues/",
            bug_link_text: "Report an issue",
            adjust_window_size: "2",
            show_close_button: "Close",
          }),
      },
      {
        label: "Check for Updates",
        click: async () => {
          axios
            .get(
              "https://api.github.com/repos/agam778/MS-365-Electron/releases/latest"
            )
            .then((res) => {
              let data = res.data;
              let currentVersion = "v" + app.getVersion();
              let latestVersion = data.tag_name;
              if (currentVersion !== latestVersion) {
                const updatedialog = dialog.showMessageBoxSync({
                  type: "info",
                  title: "Update Available",
                  message: `Your App's version: ${currentVersion}\nLatest version: ${latestVersion}\n\nPlease update to the latest version.`,
                  buttons: ["Download", "Close"],
                });
                if (updatedialog === 0) {
                  shell.openExternal(
                    "https://github.com/agam778/MS-365-Electron/releases/latest"
                  );
                }
              } else {
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "No Update Available",
                  message: `Your App's version: ${currentVersion}\nLatest version: ${latestVersion}\n\nYou are already using the latest version.`,
                  buttons: ["OK"],
                });
              }
            });
        },
      },
      {
        label: "Learn More",
        click: async () => {
          const { shell } = require("electron");
          await shell.openExternal(
            "https://github.com/agam778/MS-365-Electron"
          );
        },
      },
      {
        label: "Open Logs Folder",
        click: async () => {
          const { shell } = require("electron");
          if (process.platform === "win32") {
            await shell.openPath(
              "C:\\Users\\" +
                process.env.USERNAME +
                "\\AppData\\Roaming\\ms-365-electron\\logs\\"
            );
          } else if (process.platform === "darwin") {
            await shell.openPath(
              "/Users/" + process.env.USER + "/Library/Logs/ms-365-electron/"
            );
          } else if (process.platform === "linux") {
            await shell.openPath(
              "/home/" + process.env.USER + "/.config/ms-365-electron/logs/"
            );
          }
        },
      },
      { type: "separator" },
      {
        label: "Open Normal version of MS 365",
        type: "radio",
        click() {
          store.set("enterprise-or-normal", "https://microsoft365.com/?auth=1");
          dialog.showMessageBoxSync({
            type: "info",
            title: "Normal version of MS 365",
            message:
              "The normal version of MS 365 will be opened.\n\nPlease restart the app to apply the changes.",
            buttons: ["OK"],
          });
        },
        checked:
          store.get("enterprise-or-normal") ===
          "https://microsoft365.com/?auth=1",
      },
      {
        label: "Open Enterprise version of MS 365",
        type: "radio",
        click() {
          store.set("enterprise-or-normal", "https://microsoft365.com/?auth=2");
          dialog.showMessageBoxSync({
            type: "info",
            title: "Enterprise version of MS 365",
            message:
              "The enterprise version of MS 365 will be opened.\n\nPlease restart the app to apply the changes.",
            buttons: ["OK"],
          });
        },
        checked:
          store.get("enterprise-or-normal") ===
          "https://microsoft365.com/?auth=2",
      },
      { type: "separator" },
      {
        label: "Open Websites in New Windows (Recommended)",
        type: "radio",
        click: () => {
          store.set("websites-in-new-window", "true");
          dialog.showMessageBoxSync({
            type: "info",
            title: "Websites in New Windows",
            message:
              "Websites which are targeted to open in new tabs will now open in new windows.",
            buttons: ["OK"],
          });
        },
        checked: store.get("websites-in-new-window")
          ? store.get("websites-in-new-window") === "true"
          : true,
      },
      {
        label: "Open Websites in the Same Window",
        type: "radio",
        click: () => {
          store.set("websites-in-new-window", "false");
          dialog.showMessageBoxSync({
            type: "info",
            title: "Websites in New Windows",
            message:
              "Websites which are targeted to open in new tabs will now open in the same window.\n\nNote: This will be buggy in some cases if you are using Enterprise version of MS 365.",
            buttons: ["OK"],
          });
        },
        checked: store.get("websites-in-new-window")
          ? store.get("websites-in-new-window") === "false"
          : false,
      },
      { type: "separator" },
      {
        label: "Enable Discord Rich Presence",
        type: "checkbox",
        click: () => {
          if (store.get("discordrpcstatus") === "true") {
            store.set("discordrpcstatus", "false");
            dialog.showMessageBoxSync({
              type: "info",
              title: "Discord Rich Presence",
              message: "Discord Rich Presence is now disabled.",
              buttons: ["OK"],
            });
            rpc.clearActivity();
            return;
          } else if (
            store.get("discordrpcstatus") === "false" ||
            store.get("discordrpcstatus") === undefined
          ) {
            store.set("discordrpcstatus", "true");
            dialog.showMessageBoxSync({
              type: "info",
              title: "Discord Rich Presence",
              message: "Discord Rich Presence is now enabled.\nPlease restart the app to enable RPC.",
              buttons: ["OK"],
            });
            discordrpcupdate(
              `On "${BrowserWindow.getFocusedWindow().webContents.getTitle()}"`
            );
            return;
          }
        },
        checked: store.get("discordrpcstatus") === "true" ? true : false,
      },
      { type: "separator" },
      {
        label: "Windows User Agent String",
        type: "radio",
        click: () => {
          store.set("useragentstring", windowsuseragent);
          dialog.showMessageBoxSync({
            type: "info",
            title: "User agent switcher",
            message:
              "You have switched to Windows Useragent.\n\nPlease restart the app to apply the changes.",
            buttons: ["OK"],
          });
        },
        checked:
          store.get("useragentstring") === windowsuseragent ? true : false,
      },
      {
        label: "macOS User Agent String",
        type: "radio",
        click: () => {
          store.set("useragentstring", macuseragent);
          dialog.showMessageBoxSync({
            type: "info",
            title: "User agent switcher",
            message:
              "You have switched to Mac OS Useragent.\n\nPlease restart the app to apply the changes.",
            buttons: ["OK"],
          });
        },
        checked: store.get("useragentstring") === macuseragent ? true : false,
      },
      {
        label: "Linux User Agent String",
        type: "radio",
        click: () => {
          store.set("useragentstring", linuxuseragent);
          dialog.showMessageBoxSync({
            type: "info",
            title: "User agent switcher",
            message:
              "You have switched to Linux Useragent.\n\nPlease restart the app to apply the changes.",
            buttons: ["OK"],
          });
        },
        checked: store.get("useragentstring") === linuxuseragent ? true : false,
      },
      { type: "separator" },
      {
        role: "quit",
        accelerator: process.platform === "darwin" ? "Ctrl+Q" : "Ctrl+Q",
      },
    ],
  },
  {
    label: "Navigation",
    submenu: [
      {
        label: "Back",
        click: () => {
          BrowserWindow.getFocusedWindow().webContents.goBack();
        },
      },
      {
        label: "Forward",
        click: () => {
          BrowserWindow.getFocusedWindow().webContents.goForward();
        },
      },
      {
        label: "Reload",
        click: () => {
          BrowserWindow.getFocusedWindow().webContents.reload();
        },
      },
      {
        label: "Home",
        click: () => {
          BrowserWindow.getFocusedWindow().loadURL(
            `${store.get("enterprise-or-normal")}`
          );
        },
      },
    ],
  },
  {
    label: "Edit",
    submenu: [
      { role: "undo" },
      { role: "redo" },
      { type: "separator" },
      { role: "cut" },
      { role: "copy" },
      { role: "paste" },
      ...(process.platform === "darwin"
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
  {
    label: "Window",
    submenu: [
      { role: "minimize" },
      { role: "zoom" },
      ...(process.platform === "darwin"
        ? [
            { type: "separator" },
            { role: "front" },
            { type: "separator" },
            { role: "window" },
          ]
        : [{ role: "close" }]),
      {
        label: "Show Menu Bar",
        type: "radio",
        click: () => {
          store.set("autohide-menubar", "false");
          dialog.showMessageBoxSync({
            type: "info",
            title: "Menu Bar Settings",
            message:
              "Menu will be visible now. Please restart the app for changes to take effect.",
            buttons: ["OK"],
          });
        },
        checked: store.get("autohide-menubar") === "false",
      },
      {
        label: "Hide Menu Bar (Press ALT To show for some time)",
        type: "radio",
        click: () => {
          store.set("autohide-menubar", "true");
          dialog.showMessageBoxSync({
            type: "info",
            title: "Menu Bar Settings",
            message:
              "Menu bar will be automatically hidden now. Please restart the app for changes to take effect.",
            buttons: ["OK"],
          });
        },
        checked: store.get("autohide-menubar") === "true",
      },
    ],
  },
];

const menu = Menu.buildFromTemplate(menulayout);
Menu.setApplicationMenu(menu);

function discordrpc(title) {
  if (store.get("discordrpcstatus") === "true") {
    rpc
      .setActivity({
        details: `${title}`,
        largeImageKey: "logo",
        largeImageText: "MS-365-Electron",
        startTimestamp: Date.now(),
        instance: false,
      })
      .catch((err) => {
        console.log(err);
      });
  }
}

function discordrpcupdate(title) {
  rpc.clearActivity();
  rpc
    .setActivity({
      details: `${title}`,
      largeImageKey: "logo",
      largeImageText: "MS-365-Electron",
      startTimestamp: Date.now(),
      instance: false,
    })
    .catch((err) => {
      console.log(err);
    });
}

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
      userAgent: store.get("useragentstring") || windowsuseragent,
    }
  );

  win.webContents.on("did-finish-load", () => {
    splash.destroy();
    win.show();
    if (store.get("discordrpcstatus") === "true") {
      discordrpc(`On "${win.webContents.getTitle()}"`);
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
      }
      else {
        BrowserWindow.getFocusedWindow().loadURL(url);
        return { action: "deny" };
      }
    } else {
      return { action: "allow" };
    }
  });
});

app.on("window-all-closed", () => {
  if (store.get("discordrpcstatus") === "true") {
    rpc.clearActivity();
  }
  if (process.platform !== "darwin") {
    app.quit();
  }
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
    rpc.login({ clientId }).catch(() =>
      dialog.showMessageBox(BrowserWindow.getFocusedWindow(), {
        type: "error",
        title: "Discord RPC Error",
        message:
          "Oops! An Error occured while connecting to Discord RPC. Probably discord isn't installed or opened?",
        buttons: ["OK"],
      })
    );
  }
});
