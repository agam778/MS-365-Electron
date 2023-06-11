const { Client } = require("@xhayper/discord-rpc");
const { dialog, BrowserWindow } = require("electron");
const { setValue } = require("./store");

const client = new Client({
  clientId: "942637872530460742",
});

async function rpcError(status) {
  const rpcerror = dialog.showMessageBoxSync(BrowserWindow.getFocusedWindow(), {
    type: "error",
    title: "Discord RPC Error",
    message: `Oops! An Error occured while ${status} Discord RPC.`,
    buttons: ["Close", "Disable Discord RPC"],
  });

  if (rpcerror === 1) {
    setValue("discordrpcstatus", "false");
  }
}

async function clearActivity() {
  await client.user?.clearActivity().catch((err) => {
    rpcError("clearing");
    console.error(err);
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
      rpcError("setting");
      console.error(err);
    });
}

async function loginToRPC() {
  await client.login().catch((err) => {
    rpcError("logging into");
    console.error(err);
  });
}

module.exports = {
  clearActivity,
  setActivity,
  loginToRPC,
};
