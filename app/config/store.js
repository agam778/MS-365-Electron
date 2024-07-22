import useragents from "../useragents.json" with { type: "json" };
import Store from "electron-store";

const store = new Store();

export function getValue(key) {
  return store.get(key);
}

export function setValue(key, value) {
  store.set(key, value);
}

export function getValueOrDefault(key, defaultValue) {
  const value = store.get(key);
  if (value === undefined) {
    store.set(key, defaultValue);
    return defaultValue;
  }
  return value;
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
getValueOrDefault("windowWidth", 0.71);
getValueOrDefault("windowHeight", 0.74);
getValueOrDefault("customWindowSize", false);

if (getValue("enterprise-or-normal") === "https://microsoft365.com/?auth=1") {
  setValue("enterprise-or-normal", "?auth=1");
} else if (getValue("enterprise-or-normal") === "https://microsoft365.com/?auth=2") {
  setValue("enterprise-or-normal", "?auth=2");
}
