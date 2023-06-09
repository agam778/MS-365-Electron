const { Client } = require("@xhayper/discord-rpc");
const { dialog, BrowserWindow } = require("electron");

const client = new Client({
  clientId: "942637872530460742",
});

async function clearActivity() {
  await client.user?.clearActivity().catch((err) => {
    dialog.showMessageBox(BrowserWindow.getFocusedWindow(), {
      type: "error",
      title: "Discord RPC Error",
      message: "Oops! An Error occured while clearing Discord RPC.",
      buttons: ["OK"],
    });
  });
}

async function setActivity(details) {
  if (!client.user) {
    await loginToRPC();
  }
  await client.user
    ?.setActivity({
      details: details,
      startTimestamp: Date.now(),
      largeImageKey: "icon",
      largeImageText: "MS-365-Electron",
    })
    .catch((err) => {
      dialog.showMessageBox(BrowserWindow.getFocusedWindow(), {
        type: "error",
        title: "Discord RPC Error",
        message: "Oops! An Error occured while setting Discord RPC.",
        buttons: ["OK"],
      });
      console.error(err);
    });
}

async function loginToRPC() {
  await client.login().catch((err) => {
    dialog.showMessageBox(BrowserWindow.getFocusedWindow(), {
      type: "error",
      title: "Discord RPC Error",
      message: "Oops! An Error occured while connecting to Discord RPC.",
      buttons: ["OK"],
    });
    console.error(err);
  });
}

module.exports = {
  clearActivity,
  setActivity,
  loginToRPC,
};
