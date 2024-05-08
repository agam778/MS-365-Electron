import { app, dialog, BrowserWindow, ShareMenu, clipboard } from "electron";
import { getValue, setValue, getValueOrDefault } from './store.js';
import { ElectronBlocker } from "@cliqz/adblocker-electron";
import { clearActivity, setActivity } from './rpc.js';
import { fileURLToPath } from 'url';
import { shell } from "electron";
import { dirname } from 'path';

import useragents from "./useragents.json" with { type: "json" }
import updaterpkg from "electron-updater";
import fetch from "cross-fetch";
import axios from "axios";
import path from "path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const { autoUpdater } = updaterpkg;


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
  await shell.openExternal(url);
}

async function openLogsFolder() {
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
  setValue("useragentstring", useragent);
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

getValueOrDefault("enterprise-or-normal", "?auth=1");
getValueOrDefault("websites-in-new-window", "true");
getValueOrDefault("autohide-menubar", "false");
getValueOrDefault("useragentstring", useragents.Windows);
getValueOrDefault("discordrpcstatus", "false");
getValueOrDefault("blockadsandtrackers", "false");
getValueOrDefault("dynamicicons", "true");
getValueOrDefault("autoupdater", "true");
getValueOrDefault("custompage", "home");

// temporary code to change the value of enterprise-or-normal to the new values for the update, will be removed in the next update.
if (getValue("enterprise-or-normal") === "https://microsoft365.com/?auth=1") {
  setValue("enterprise-or-normal", "?auth=1");
} else if (
  getValue("enterprise-or-normal") === "https://microsoft365.com/?auth=2"
) {
  setValue("enterprise-or-normal", "?auth=2");
}

const menulayout = [
  ...(process.platform === "darwin"
    ? [
        {
          label: app.name,
          submenu: [
            {
              label: "About MS-365-Electron",
              click: async () => {
                await openExternalLink(
                  "https://github.com/agam778/MS-365-Electron"
                );
              },
            },
            {
              label: "Check for Updates",
              id: "check-for-updates",
              click: async () => {
                await checkForUpdates();
              },
            },
            { type: "separator" },
            {
              label: "Open Logs Folder",
              click: async () => {
                await openLogsFolder();
              },
            },
            { type: "separator" },
            {
              label: "Preferences",
              submenu: [
                {
                  label: "Open MS 365 with Personal Account",
                  type: "radio",
                  click() {
                    setValue(
                      "enterprise-or-normal",
                      "?auth=1"
                    );
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "MS 365 with Personal Account",
                      message:
                        "MS 365 will now open with your Personal Account.\n\nPlease restart the app to apply the changes.",
                      buttons: ["OK"],
                    });
                  },
                  checked:
                    getValue("enterprise-or-normal") ===
                    "?auth=1",
                },
                {
                  label: "Open MS 365 with Work/School Account",
                  type: "radio",
                  click() {
                    setValue(
                      "enterprise-or-normal",
                      "?auth=2"
                    );
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "MS 365 with Work/School Account",
                      message:
                        "MS 365 will now open with your Work/School account.\n\nPlease restart the app to apply the changes.",
                      buttons: ["OK"],
                    });
                  },
                  checked:
                    getValue("enterprise-or-normal") ===
                    "?auth=2",
                },
                { type: "separator" },
                {
                  label: "Open Websites in New Windows",
                  type: "radio",
                  click: () => {
                    setValue("websites-in-new-window", "true");
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "Websites in New Windows",
                      message:
                        "Websites which are targeted to open in new tabs will now open in new windows.",
                      buttons: ["OK"],
                    });
                  },
                  checked: getValue("websites-in-new-window") === "true",
                },
                {
                  label: "Open Websites in the Same Window",
                  type: "radio",
                  click: () => {
                    setValue("websites-in-new-window", "false");
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "Websites in New Windows",
                      message:
                        "Websites which are targeted to open in new tabs will now open in the same window.",
                      buttons: ["OK"],
                    });
                  },
                  checked: getValue("websites-in-new-window") === "false",
                },
                { type: "separator" },
                {
                  label: "Custom Home Page",
                  submenu: [
                    {
                      label: "Home (Default)",
                      type: "radio",
                      click: () => {
                        setValue("custompage", "home");
                        dialog.showMessageBoxSync({
                          type: "info",
                          title: "Custom Home Page",
                          message:
                            "You have set the home page to the default home page. Please restart the app to apply the changes.",
                          buttons: ["OK"],
                        });
                      },
                      checked: getValue("custompage") === "home",
                    },
                    {
                      label: "Create",
                      type: "radio",
                      click: () => {
                        setValue("custompage", "create");
                        dialog.showMessageBoxSync({
                          type: "info",
                          title: "Custom Home Page",
                          message:
                            "You have set the home page to the \"Create\" page. Please restart the app to apply the changes.",
                          buttons: ["OK"],
                        });
                      },
                      checked: getValue("custompage") === "create",
                    },
                    {
                      label: "My Content",
                      type: "radio",
                      click: () => {
                        setValue("custompage", "mycontent");
                        dialog.showMessageBoxSync({
                          type: "info",
                          title: "Custom Home Page",
                          message:
                            "You have set the home page to the \"My Content\" page. Please restart the app to apply the changes.",
                          buttons: ["OK"],
                        });
                      },
                      checked: getValue("custompage") === "mycontent",
                    },
                    {
                      label: "Apps",
                      type: "radio",
                      click: () => {
                        setValue("custompage", "apps");
                        dialog.showMessageBoxSync({
                          type: "info",
                          title: "Custom Home Page",
                          message:
                            "You have set the home page to the \"Apps\" page. Please restart the app to apply the changes.",
                          buttons: ["OK"],
                        });
                      },
                      checked: getValue("custompage") === "apps",
                    },
                  ],
                },
                { type: "separator" },
                {
                  label: "Enable Discord RPC",
                  type: "checkbox",
                  click: () => {
                    if (getValue("discordrpcstatus") === "true") {
                      setValue("discordrpcstatus", "false");
                      dialog.showMessageBoxSync({
                        type: "info",
                        title: "Discord RPC",
                        message: "Discord RPC has been disabled.",
                        buttons: ["OK"],
                      });
                      clearActivity();
                      return;
                    } else if (
                      getValue("discordrpcstatus") === "false" ||
                      getValue("discordrpcstatus") === undefined
                    ) {
                      setValue("discordrpcstatus", "true");
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
                  checked: getValue("discordrpcstatus") === "true",
                },
                {
                  label: "Enable Auto Updates",
                  type: "checkbox",
                  click: () => {
                    if (getValue("autoupdater") === "true") {
                      setValue("autoupdater", "false");
                      dialog.showMessageBoxSync({
                        type: "info",
                        title: "Auto Updates",
                        message: "Auto updates have been disabled.",
                        buttons: ["OK"],
                      });
                      return;
                    } else if (
                      getValue("autoupdater") === "false" ||
                      getValue("autoupdater") === undefined
                    ) {
                      setValue("autoupdater", "true");
                      dialog.showMessageBoxSync({
                        type: "info",
                        title: "Auto Updates",
                        message: "Auto updates have been enabled.",
                        buttons: ["OK"],
                      });
                      return;
                    }
                  },
                  checked: getValue("autoupdater") === "true",
                },
                {
                  label: "Enable Dynamic Icons",
                  type: "checkbox",
                  click: () => {
                    if (getValue("dynamicicons") === "true") {
                      setValue("dynamicicons", "false");
                      dialog.showMessageBoxSync({
                        type: "info",
                        title: "Dynamic Icons",
                        message: "Dynamic icons have been disabled.",
                        buttons: ["OK"],
                      });
                      return;
                    } else if (
                      getValue("dynamicicons") === "false" ||
                      getValue("dynamicicons") === undefined
                    ) {
                      setValue("dynamicicons", "true");
                      dialog.showMessageBoxSync({
                        type: "info",
                        title: "Dynamic Icons",
                        message: "Dynamic icons have been enabled.",
                        buttons: ["OK"],
                      });
                      return;
                    }
                  },
                  checked: getValue("dynamicicons") === "true",
                },
                { type: "separator" },
                {
                  label: "Block Ads and Trackers",
                  type: "checkbox",
                  click: () => {
                    if (getValue("blockadsandtrackers") === "true") {
                      setValue("blockadsandtrackers", "false");
                      dialog.showMessageBoxSync({
                        type: "info",
                        title: "Block Ads and Trackers",
                        message: "Ads and trackers will no longer be blocked.",
                        buttons: ["OK"],
                      });
                      ElectronBlocker.fromPrebuiltAdsAndTracking(fetch).then(
                        (blocker) => {
                          blocker.disableBlockingInSession(
                            BrowserWindow.getFocusedWindow().webContents.session
                          );
                        }
                      );
                      return;
                    }
                    if (
                      getValue("blockadsandtrackers") === "false" ||
                      getValue("blockadsandtrackers") === undefined
                    ) {
                      setValue("blockadsandtrackers", "true");
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
                  checked: getValue("blockadsandtrackers") === "true",
                },
                { type: "separator" },
                {
                  label: "Windows User-Agent String",
                  type: "radio",
                  click: () => {
                    setUserAgent(useragents.Windows);
                  },
                  checked: getValue("useragentstring") === useragents.Windows,
                },
                {
                  label: "macOS User-Agent String",
                  type: "radio",
                  click: () => {
                    setUserAgent(useragents.macOS);
                  },
                  checked: getValue("useragentstring") === useragents.macOS,
                },
                {
                  label: "Linux User-Agent String",
                  type: "radio",
                  click: () => {
                    setValue("useragentstring", useragents.Linux);
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "User agent switcher",
                      message:
                        "You have switched to Linux Useragent.\n\nPlease restart the app to apply the changes.",
                      buttons: ["OK"],
                    });
                  },
                  checked: getValue("useragentstring") === useragents.Linux,
                },
                { type: "separator" },
                ...(process.platform === "win32" || process.platform === "linux"
                  ? [
                      {
                        role: "quit",
                        accelerator: "Ctrl+Q",
                      },
                    ]
                  : []),
              ],
            },
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
  ...(process.platform === "win32" || process.platform === "linux"
    ? [
        {
          label: "MS-365-Electron",
          submenu: [
            {
              label: "About MS-365-Electron",
              click: async () => {
                await openExternalLink(
                  "https://github.com/agam778/MS-365-Electron"
                );
              },
            },
            {
              label: "Check for Updates...",
              click: async () => {
                await checkForUpdates();
              },
            },
            { type: "separator" },
            {
              label: "Open Logs Folder",
              click: async () => {
                await openLogsFolder();
              },
            },
            { type: "separator" },
            {
              label: "Open MS 365 with Personal Account",
              type: "radio",
              click() {
                setValue(
                  "enterprise-or-normal",
                  "?auth=1"
                );
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "MS 365 with Personal Account",
                  message:
                    "MS 365 will now open with your Personal Account.\n\nPlease restart the app to apply the changes.",
                  buttons: ["OK"],
                });
              },
              checked:
                getValue("enterprise-or-normal") ===
                "?auth=1",
            },
            {
              label: "Open MS 365 with Work/School Account",
              type: "radio",
              click() {
                setValue(
                  "enterprise-or-normal",
                  "?auth=2"
                );
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "MS 365 with Work/School Account",
                  message:
                    "MS 365 will now open with your Work/School account.\n\nPlease restart the app to apply the changes.",
                  buttons: ["OK"],
                });
              },
              checked:
                getValue("enterprise-or-normal") ===
                "?auth=2",
            },
            { type: "separator" },
            {
              label: "Open Websites in New Windows",
              type: "radio",
              click: () => {
                setValue("websites-in-new-window", "true");
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "Websites in New Windows",
                  message:
                    "Websites which are targeted to open in new tabs will now open in new windows.",
                  buttons: ["OK"],
                });
              },
              checked: getValue("websites-in-new-window") === "true",
            },
            {
              label: "Open Websites in the Same Window",
              type: "radio",
              click: () => {
                setValue("websites-in-new-window", "false");
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "Websites in New Windows",
                  message:
                    "Websites which are targeted to open in new tabs will now open in the same window.",
                  buttons: ["OK"],
                });
              },
              checked: getValue("websites-in-new-window") === "false",
            },
            { type: "separator" },
            {
              label: "Custom Home Page",
              submenu: [
                {
                  label: "Home (Default)",
                  type: "radio",
                  click: () => {
                    setValue("custompage", "home");
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "Custom Home Page",
                      message:
                        "You have set the home page to the default home page. Please restart the app to apply the changes.",
                      buttons: ["OK"],
                    });
                  },
                  checked: getValue("custompage") === "home",
                },
                {
                  label: "Create",
                  type: "radio",
                  click: () => {
                    setValue("custompage", "create");
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "Custom Home Page",
                      message:
                        "You have set the home page to the \"Create\" page. Please restart the app to apply the changes.",
                      buttons: ["OK"],
                    });
                  },
                  checked: getValue("custompage") === "create",
                },
                {
                  label: "My Content",
                  type: "radio",
                  click: () => {
                    setValue("custompage", "mycontent");
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "Custom Home Page",
                      message:
                        "You have set the home page to the \"My Content\" page. Please restart the app to apply the changes.",
                      buttons: ["OK"],
                    });
                  },
                  checked: getValue("custompage") === "mycontent",
                },
                {
                  label: "Apps",
                  type: "radio",
                  click: () => {
                    setValue("custompage", "apps");
                    dialog.showMessageBoxSync({
                      type: "info",
                      title: "Custom Home Page",
                      message:
                        "You have set the home page to the \"Apps\" page. Please restart the app to apply the changes.",
                      buttons: ["OK"],
                    });
                  },
                  checked: getValue("custompage") === "apps",
                },
              ],
            },
            { type: "separator" },
            {
              label: "Enable Discord RPC",
              type: "checkbox",
              click: () => {
                if (getValue("discordrpcstatus") === "true") {
                  setValue("discordrpcstatus", "false");
                  dialog.showMessageBoxSync({
                    type: "info",
                    title: "Discord RPC",
                    message: "Discord RPC has been disabled.",
                    buttons: ["OK"],
                  });
                  clearActivity();
                  return;
                } else if (
                  getValue("discordrpcstatus") === "false" ||
                  getValue("discordrpcstatus") === undefined
                ) {
                  setValue("discordrpcstatus", "true");
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
              checked: getValue("discordrpcstatus") === "true",
            },
            {
              label: "Enable Auto Updates",
              type: "checkbox",
              click: () => {
                if (getValue("autoupdater") === "true") {
                  setValue("autoupdater", "false");
                  dialog.showMessageBoxSync({
                    type: "info",
                    title: "Auto Updates",
                    message: "Auto updates have been disabled.",
                    buttons: ["OK"],
                  });
                  return;
                } else if (
                  getValue("autoupdater") === "false" ||
                  getValue("autoupdater") === undefined
                ) {
                  setValue("autoupdater", "true");
                  dialog.showMessageBoxSync({
                    type: "info",
                    title: "Auto Updates",
                    message: "Auto updates have been enabled.",
                    buttons: ["OK"],
                  });
                  return;
                }
              },
              checked: getValue("autoupdater") === "true",
            },
            {
              label: "Enable Dynamic Icons",
              type: "checkbox",
              click: () => {
                if (getValue("dynamicicons") === "true") {
                  setValue("dynamicicons", "false");
                  dialog.showMessageBoxSync({
                    type: "info",
                    title: "Dynamic Icons",
                    message: "Dynamic icons have been disabled.",
                    buttons: ["OK"],
                  });
                  return;
                } else if (
                  getValue("dynamicicons") === "false" ||
                  getValue("dynamicicons") === undefined
                ) {
                  setValue("dynamicicons", "true");
                  dialog.showMessageBoxSync({
                    type: "info",
                    title: "Dynamic Icons",
                    message: "Dynamic icons have been enabled.",
                    buttons: ["OK"],
                  });
                  return;
                }
              },
              checked: getValue("dynamicicons") === "true",
            },
            { type: "separator" },
            {
              label: "Block Ads and Trackers",
              type: "checkbox",
              click: () => {
                if (getValue("blockadsandtrackers") === "true") {
                  setValue("blockadsandtrackers", "false");
                  dialog.showMessageBoxSync({
                    type: "info",
                    title: "Block Ads and Trackers",
                    message: "Ads and trackers will no longer be blocked.",
                    buttons: ["OK"],
                  });
                  ElectronBlocker.fromPrebuiltAdsAndTracking(fetch).then(
                    (blocker) => {
                      blocker.disableBlockingInSession(
                        BrowserWindow.getFocusedWindow().webContents.session
                      );
                    }
                  );
                  return;
                }
                if (
                  getValue("blockadsandtrackers") === "false" ||
                  getValue("blockadsandtrackers") === undefined
                ) {
                  setValue("blockadsandtrackers", "true");
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
              checked: getValue("blockadsandtrackers") === "true",
            },
            { type: "separator" },
            {
              label: "Windows User-Agent String",
              type: "radio",
              click: () => {
                setUserAgent(useragents.Windows);
              },
              checked: getValue("useragentstring") === useragents.Windows,
            },
            {
              label: "macOS User-Agent String",
              type: "radio",
              click: () => {
                setUserAgent(useragents.macOS);
              },
              checked: getValue("useragentstring") === useragents.macOS,
            },
            {
              label: "Linux User-Agent String",
              type: "radio",
              click: () => {
                setValue("useragentstring", useragents.Linux);
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "User agent switcher",
                  message:
                    "You have switched to Linux Useragent.\n\nPlease restart the app to apply the changes.",
                  buttons: ["OK"],
                });
              },
              checked: getValue("useragentstring") === useragents.Linux,
            },
            { type: "separator" },
            ...(process.platform === "win32" || process.platform === "linux"
              ? [
                  {
                    role: "quit",
                    accelerator: "Ctrl+Q",
                  },
                ]
              : []),
          ],
        },
      ]
    : []),
  {
    label: "File",
    submenu: [
      {
        label: "New Window (Personal)",
        accelerator: "CmdOrCtrl+N",
        click: () => {
          let newWindow = new BrowserWindow({
            width: 1081,
            height: 570,
            webPreferences: {
              nodeIntegration: true,
              devTools: true,
              partition: "persist:personal",
            },
          });
          newWindow.loadURL(`https://microsoft365.com/${getValue("custompage")}/?auth=1`);
        },
      },
      {
        label: "New Window (Work/School)",
        accelerator: "CmdOrCtrl+Shift+N",
        click: () => {
          let newWindow = new BrowserWindow({
            width: 1081,
            height: 570,
            webPreferences: {
              nodeIntegration: true,
              devTools: true,
              partition: "persist:work",
            },
          });
          newWindow.loadURL(`https://microsoft365.com/${getValue("custompage")}/?auth=2`);
        },
      },
      { type: "separator" },
      {
        label: "Close Window",
        accelerator: "CmdOrCtrl+W",
        click: () => {
          try {
            BrowserWindow.getFocusedWindow().close();
          } catch {
            return;
          }
        },
      },
      {
        label: "Close All Windows",
        accelerator: "CmdOrCtrl+Shift+W",
        click: () => {
          BrowserWindow.getAllWindows().forEach((window) => {
            window.close();
          });
        },
      },
      { type: "separator" },
      {
        label: "Copy URL to Clipboard",
        accelerator: "CmdOrCtrl+Shift+C",
        click: () => {
          const url = BrowserWindow.getFocusedWindow().webContents.getURL();
          clipboard.writeText(url);
        },
      },
      ...(process.platform === "darwin"
        ? [
            {
              label: "Share...",
              click: () => {
                let sharemenu = new ShareMenu({
                  urls: [BrowserWindow.getFocusedWindow().webContents.getURL()],
                  texts: [BrowserWindow.getFocusedWindow().getTitle()],
                });
                sharemenu.popup();
              },
            },
          ]
        : []),
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
    label: "Navigation",
    submenu: [
      {
        label: "Back",
        click: () => {
          BrowserWindow.getFocusedWindow().webContents.goBack();
        },
        accelerator: "Alt+Left",
      },
      {
        label: "Forward",
        click: () => {
          BrowserWindow.getFocusedWindow().webContents.goForward();
        },
        accelerator: "Alt+Right",
      },
      {
        label: "Reload",
        click: () => {
          BrowserWindow.getFocusedWindow().webContents.reload();
        },
        accelerator: "CmdOrCtrl+R",
      },
      {
        label: "Home",
        click: () => {
          BrowserWindow.getFocusedWindow().loadURL(
            `${getValue("enterprise-or-normal")}`
          );
        },
      },
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
    label: "Apps",
    submenu: [
      {
        label: "Word",
        click: () => {
          if (
            getValue("enterprise-or-normal") ===
            "?auth=2"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let wordwindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:work",
                },
              });
              wordwindow.loadURL("https://microsoft365.com/launch/word?auth=2");
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://microsoft365.com/launch/word?auth=2"
              );
            }
          } else if (
            getValue("enterprise-or-normal") ===
            "?auth=1"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let wordwindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:personal",
                },
              });
              wordwindow.loadURL("https://microsoft365.com/launch/word?auth=1");
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://microsoft365.com/launch/word?auth=1"
              );
            }
          }
        },
      },
      {
        label: "Excel",
        click: () => {
          if (
            getValue("enterprise-or-normal") ===
            "?auth=2"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let excelwindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:work",
                },
              });
              excelwindow.loadURL(
                "https://microsoft365.com/launch/excel?auth=2"
              );
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://microsoft365.com/launch/excel?auth=2"
              );
            }
          } else if (
            getValue("enterprise-or-normal") ===
            "?auth=1"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let excelwindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:personal",
                },
              });
              excelwindow.loadURL(
                "https://microsoft365.com/launch/excel?auth=1"
              );
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://microsoft365.com/launch/excel?auth=1"
              );
            }
          }
        },
      },
      {
        label: "PowerPoint",
        click: () => {
          if (
            getValue("enterprise-or-normal") ===
            "?auth=2"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let powerpointwindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:work",
                },
              });
              powerpointwindow.loadURL(
                "https://microsoft365.com/launch/powerpoint?auth=2"
              );
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://microsoft365.com/launch/powerpoint?auth=2"
              );
            }
          } else if (
            getValue("enterprise-or-normal") ===
            "?auth=1"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let powerpointwindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:personal",
                },
              });
              powerpointwindow.loadURL(
                "https://microsoft365.com/launch/powerpoint?auth=1"
              );
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://microsoft365.com/launch/powerpoint?auth=1"
              );
            }
          }
        },
      },
      {
        label: "Outlook",
        click: () => {
          if (
            getValue("enterprise-or-normal") ===
            "?auth=2"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let outlookwindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:work",
                },
              });
              outlookwindow.loadURL("https://outlook.office.com/mail/");
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://outlook.office.com/mail/"
              );
            }
          } else if (
            getValue("enterprise-or-normal") ===
            "?auth=1"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let outlookwindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:personal",
                },
              });
              outlookwindow.loadURL(
                "https://office.live.com/start/Outlook.aspx"
              );
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://office.live.com/start/Outlook.aspx"
              );
            }
          }
        },
      },
      {
        label: "OneDrive",
        click: () => {
          if (
            getValue("enterprise-or-normal") ===
            "?auth=2"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let onedrivewindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:work",
                },
              });
              onedrivewindow.loadURL(
                "https://microsoft365.com/launch/onedrive?auth=2"
              );
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://microsoft365.com/launch/onedrive?auth=2"
              );
            }
          } else if (
            getValue("enterprise-or-normal") ===
            "?auth=1"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let onedrivewindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:personal",
                },
              });
              onedrivewindow.loadURL(
                "https://microsoft365.com/launch/onedrive?auth=1"
              );
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://microsoft365.com/launch/onedrive?auth=1"
              );
            }
          }
        },
      },
      {
        label: "OneNote",
        click: () => {
          if (
            getValue("enterprise-or-normal") ===
            "?auth=2"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let onenotewindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:work",
                },
              });
              onenotewindow.loadURL(
                "https://www.microsoft365.com/launch/onenote?auth=2"
              );
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://www.microsoft365.com/launch/onenote?auth=2"
              );
            }
          } else if (
            getValue("enterprise-or-normal") ===
            "?auth=1"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let onenotewindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:personal",
                },
              });
              onenotewindow.loadURL("https://www.onenote.com/notebooks?auth=1");
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://www.onenote.com/notebooks?auth=1"
              );
            }
          }
        },
      },
      {
        label: "All Apps",
        click: () => {
          if (
            getValue("enterprise-or-normal") ===
            "?auth=2"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let allappswindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:work",
                },
              });
              allappswindow.loadURL("https://www.microsoft365.com/apps?auth=2");
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://www.microsoft365.com/apps?auth=2"
              );
            }
          } else if (
            getValue("enterprise-or-normal") ===
            "?auth=1"
          ) {
            if (getValue("websites-in-new-window") === "true") {
              let allappswindow = new BrowserWindow({
                width: 1081,
                height: 570,
                webPreferences: {
                  nodeIntegration: false,
                  contextIsolation: true,
                  partition: "persist:personal",
                },
              });
              allappswindow.loadURL("https://www.microsoft365.com/apps?auth=1");
            } else {
              BrowserWindow.getFocusedWindow().loadURL(
                "https://www.microsoft365.com/apps?auth=1"
              );
            }
          }
        },
      },
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
      ...(process.platform === "win32" || process.platform === "linux"
        ? [
            { type: "separator" },
            {
              label: "Show Menu Bar",
              type: "radio",
              click: () => {
                setValue("autohide-menubar", "false");
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "Menu Bar Settings",
                  message:
                    "Menu will be visible now. Please restart the app for changes to take effect.",
                  buttons: ["OK"],
                });
              },
              checked: getValue("autohide-menubar") === "false",
            },
            {
              label: "Hide Menu Bar (ALT to show)",
              type: "radio",
              click: () => {
                setValue("autohide-menubar", "true");
                dialog.showMessageBoxSync({
                  type: "info",
                  title: "Menu Bar Settings",
                  message:
                    "Menu bar will be automatically hidden now. Please restart the app for changes to take effect.",
                  buttons: ["OK"],
                });
              },
              checked: getValue("autohide-menubar") === "true",
            },
          ]
        : []),
    ],
  },
];

export default menulayout;
