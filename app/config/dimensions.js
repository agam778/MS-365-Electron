import { app, screen } from "electron";
let screenWidth, screenHeight;

app.on("ready", () => {
  ({ width: screenWidth, height: screenHeight } = screen.getPrimaryDisplay().workAreaSize);
});

export function getScreenWidth() {
  return screenWidth;
}

export function getScreenHeight() {
  return screenHeight;
}
