import { app, screen } from "electron";

let screenWidth, screenHeight;
app.on("ready", () => {
  ({ width: screenWidth, height: screenHeight } = screen.getPrimaryDisplay().workAreaSize);
});

function getScreenWidth() {
  return screenWidth;
}

function getScreenHeight() {
  return screenHeight;
}

export { getScreenWidth, getScreenHeight };
