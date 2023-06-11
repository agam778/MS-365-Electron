const Store = require("electron-store");
const store = new Store();

function getValue(key) {
  return store.get(key);
}

function setValue(key, value) {
  store.set(key, value);
}

function getValueOrDefault(key, defaultValue) {
  const value = store.get(key);
  if (value === undefined) {
    store.set(key, defaultValue);
    return defaultValue;
  }
  return value;
}

module.exports = { getValue, setValue, getValueOrDefault };
