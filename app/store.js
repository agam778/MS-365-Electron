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
