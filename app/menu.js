const store = require("./store");
const useragents = require("./useragents.json");
const { app, dialog, BrowserWindow } = require("electron");
const axios = require("axios");
const { clearActivity, setActivity } = require("./rpc");
const { shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const { ElectronBlocker } = require("@cliqz/adblocker-electron");
const fetch = require("cross-fetch");

function getValueOrDefault(key, defaultValue) {
  const value = store.get(key);
  if (value === undefined) {
    store.set(key, defaultValue);
    return defaultValue;
  }
  return value;
}

async function checkForUpdates() {
  try {
    const res = await axios.get(
      "https://api.github.com/repos/agam778/MS-365-Electron/releases/latest"
    );
    const data = res.data;
    const currentVersion = "v" + app.getVersion();
    const latestVersion = data.tag_name;

    if (currentVersion !== latestVersion) {
      if (process.platform === "win32" || process.platform === "darwin") {
        autoUpdater.checkForUpdatesAndNotify().then((result) => {
          if (result === null) {
            dialog.showMessageBoxSync({
              type: "info",
              title: "No Update Available",
              message: `Current version: ${currentVersion}\nLatest version: ${latestVersion}\n\nYou are already using the latest version.`,
              buttons: ["OK"],
            });
          }
        });
        return;
      } else {
        const updatedialog = dialog.showMessageBoxSync({
          type: "info",
          title: "Update Available",
          message: `Current version: ${currentVersion}\nLatest version: ${latestVersion}\n\nPlease update to the latest version.`,
          buttons: ["Download", "Close"],
        });
        if (updatedialog === 0) {
          shell.openExternal(
            "https://github.com/agam778/MS-365-Electron/releases/latest"
          );
        }
      }
    } else {
      dialog.showMessageBoxSync({
        type: "info",
        title: "No Update Available",
        message: `Your App's version: ${currentVersion}\nLatest version: ${latestVersion}\n\nYou are already using the latest version.`,
        buttons: ["OK"],
      });
    }
  } catch (error) {
    console.error("Error checking for updates:", error);
  }
}

async function openExternalLink(url) {
  const { shell } = require("electron");
  await shell.openExternal(url);
}

async function openLogsFolder() {
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
}

function setUserAgent(useragent) {
  store.set("useragentstring", useragent);
  const updatedialog = dialog.showMessageBoxSync({
    type: "info",
    title: "User-Agent string changed",
    message: `You have switched to the ${useragent} User-Agent string.\n\nPlease restart the app for the changes to take effect.`,
    buttons: ["Later", "Restart"],
  });
  if (updatedialog === 1) {
    app.relaunch();
    app.exit();
  }
}

getValueOrDefault("enterprise-or-normal", "https://microsoft365.com/?auth=1");
getValueOrDefault("autohide-menubar", "false");
getValueOrDefault("useragentstring", useragents.Windows);
getValueOrDefault("discordrpcstatus", "false");
getValueOrDefault("blockads", "false");
getValueOrDefault("blockadsandtrackers", "false");

const menulayout = [
  ...(process.platform === "darwin"
    ? [
        {
          label: app.name,
          submenu: [
            { label: "About MS-365-Electron", role: "about" },
            {
              label: "Learn More",
              click: async () => {
                await openExternalLink(
                  "https://github.com/agam778/MS-365-Electron"
                );
              },
            },
            {
              label: "Check for Updates...",
              id: "check-for-updates",
              click: async () => {
                await checkForUpdates();
              },
            },
            {
              label: "Open Logs Folder",
              click: async () => {
                await openLogsFolder();
              },
            },
            { type: "separator" },
            { role: "services" },
            { type: "separator" },
            { label: "Hide MS-365-Electron", role: "hide" },
            { role: "hideOthers" },
            { role: "unhide" },
            { type: "separator" },
            { label: "Quit MS-365-Electron", role: "quit" },
          ],
        },
      ]
    : []),
  {
    label: process.platform === "darwin" ? "Preferences" : "Application",
    submenu: [
      ...(!process.platform === "darwin"
        ? [
            {
              label: "About MS-365-Electron",
              click: () => {
                // placeholder
              },
            },
          ]
        : []),
      ...(!process.platform === "darwin"
        ? [
            {
              label: "Learn More",
              click: async () => {
                await openExternalLink(
                  "https://github.com/agam778/MS-365-Electron"
                );
              },
            },
          ]
        : []),
      ...(!process.platform === "darwin"
        ? [
            {
              label: "Check for Updates...",
              id: "check-for-updates",
              click: async () => {
                await checkForUpdates();
              },
            },
          ]
        : []),
      ...(!process.platform === "darwin"
        ? [
            {
              label: "Open Logs Folder",
              click: async () => {
                await openLogsFolder();
              },
            },
          ]
        : []),
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
        label: "Enable Discord RPC",
        type: "checkbox",
        click: () => {
          if (store.get("discordrpcstatus") === "true") {
            store.set("discordrpcstatus", "false");
            dialog.showMessageBoxSync({
              type: "info",
              title: "Discord RPC",
              message: "Discord RPC has been disabled.",
              buttons: ["OK"],
            });
            clearActivity();
            return;
          } else if (
            store.get("discordrpcstatus") === "false" ||
            store.get("discordrpcstatus") === undefined
          ) {
            store.set("discordrpcstatus", "true");
            dialog.showMessageBoxSync({
              type: "info",
              title: "Discord RPC",
              message: "Discord RPC has been enabled.",
              buttons: ["OK"],
            });
            setActivity(
              `On ${BrowserWindow.getFocusedWindow().webContents.getTitle()}`
            );
            return;
          }
        },
        checked: store.get("discordrpcstatus") === "true",
      },
      { type: "separator" },
      {
        label: "Block Ads",
        type: "checkbox",
        click: () => {
          if (store.get("blockads") === "true") {
            store.set("blockads", "false");
            dialog.showMessageBoxSync({
              type: "info",
              title: "Block Ads",
              message: "Ads will no longer be blocked.",
              buttons: ["OK"],
            });
            if (!store.get("blockadsandtrackers") === "true") {
              ElectronBlocker.fromPrebuiltAdsOnly(fetch).then((blocker) =>
                blocker.disableBlockingInSession(
                  BrowserWindow.getFocusedWindow().webContents.session
                )
              );
            }
            return;
          }
          if (
            store.get("blockads") === "false" ||
            store.get("blockads") === undefined
          ) {
            store.set("blockads", "true");
            ElectronBlocker.fromPrebuiltAdsOnly(fetch).then((blocker) =>
              blocker.enableBlockingInSession(
                BrowserWindow.getFocusedWindow().webContents.session
              )
            );

            dialog.showMessageBoxSync({
              type: "info",
              title: "Block Ads",
              message: "Ads will now be blocked.",
              buttons: ["OK"],
            });
            return;
          }
        },
        checked: store.get("blockads") === "true",
      },
      {
        label: "Block Ads and Trackers",
        type: "checkbox",
        click: () => {
          if (store.get("blockadsandtrackers") === "true") {
            store.set("blockadsandtrackers", "false");
            dialog.showMessageBoxSync({
              type: "info",
              title: "Block Ads and Trackers",
              message: "Ads and trackers will no longer be blocked.",
              buttons: ["OK"],
            });
            if (!store.get("blockads") === "true") {
              ElectronBlocker.fromPrebuiltAdsAndTracking(fetch).then(
                (blocker) => {
                  blocker.disableBlockingInSession(
                    BrowserWindow.getFocusedWindow().webContents.session
                  );
                }
              );
            }
            return;
          }
          if (
            store.get("blockadsandtrackers") === "false" ||
            store.get("blockadsandtrackers") === undefined
          ) {
            store.set("blockadsandtrackers", "true");
            ElectronBlocker.fromPrebuiltAdsAndTracking(fetch).then(
              (blocker) => {
                blocker.enableBlockingInSession(
                  BrowserWindow.getFocusedWindow().webContents.session
                );
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "Block Ads and Trackers",
                  message: "Ads and trackers will now be blocked.",
                  buttons: ["OK"],
                });
              }
            );
            return;
          }
        },
        checked: store.get("blockadsandtrackers") === "true",
      },
      { type: "separator" },
      {
        label: "Windows User-Agent String",
        type: "radio",
        click: () => {
          setUserAgent(useragents.Windows);
        },
        checked: store.get("useragentstring") === useragents.Windows,
      },
      {
        label: "macOS User-Agent String",
        type: "radio",
        click: () => {
          setUserAgent(useragents.macOS);
        },
        checked: store.get("useragentstring") === useragents.macOS,
      },
      {
        label: "Linux User-Agent String",
        type: "radio",
        click: () => {
          store.set("useragentstring", useragents.Linux);
          dialog.showMessageBoxSync({
            type: "info",
            title: "User agent switcher",
            message:
              "You have switched to Linux Useragent.\n\nPlease restart the app to apply the changes.",
            buttons: ["OK"],
          });
        },
        checked: store.get("useragentstring") === useragents.Linux,
      },
      { type: "separator" },
      ...(!process.platform === "darwin"
        ? [
            {
              role: "quit",
              accelerator: "Ctrl+Q",
            },
          ]
        : []),
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
        ? [{ type: "separator" }, { role: "front" }, { type: "separator" }]
        : [{ role: "close" }]),
      ...(!process.platform === "darwin"
        ? [
            { type: "separator" },
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
              label: "Hide Menu Bar (ALT to show)",
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
          ]
        : []),
    ],
  },
];
module.exports = { menulayout };
