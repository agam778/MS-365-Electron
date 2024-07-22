import { app, dialog, shell } from "electron";
import axios from "axios";
import { setValue } from "./store.js";
import updaterpkg from "electron-updater";

const { autoUpdater } = updaterpkg;

export async function checkForUpdates() {
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
          shell.openExternal("https://github.com/agam778/MS-365-Electron/releases/latest");
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

export async function openExternalLink(url) {
  await shell.openExternal(url);
}

export async function openLogsFolder() {
  if (process.platform === "win32") {
    await shell.openPath(
      "C:\\Users\\" + process.env.USERNAME + "\\AppData\\Roaming\\ms-365-electron\\logs\\"
    );
  } else if (process.platform === "darwin") {
    await shell.openPath("/Users/" + process.env.USER + "/Library/Logs/ms-365-electron/");
  } else if (process.platform === "linux") {
    await shell.openPath("/home/" + process.env.USER + "/.config/ms-365-electron/logs/");
  }
}

export function setUserAgent(useragent) {
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
