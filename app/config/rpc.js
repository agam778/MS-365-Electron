import { dialog, BrowserWindow } from "electron";
import { Client } from "@xhayper/discord-rpc";
import { setValue } from "./store.js";

const client = new Client({
  clientId: "942637872530460742",
});

export async function rpcError(status) {
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

export async function clearActivity() {
  await client.user?.clearActivity().catch((err) => {
    rpcError("clearing");
    console.error(err);
  });
}

export async function setActivity(details) {
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

export async function loginToRPC() {
  await client.login().catch((err) => {
    rpcError("logging into");
    console.error(err);
  });
}
