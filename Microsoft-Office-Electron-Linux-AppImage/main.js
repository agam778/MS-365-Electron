const { app, Menu, BrowserWindow } = require('electron')
const openAboutWindow = require('about-window').default;
const join = require('path').join;
const isMac = process.platform === 'darwin'
const path = require('path')
const electron = require('electron')
const {autoUpdater} = require("electron-updater");

const template = [
  // { role: 'appMenu' }
  ...(isMac ? [{
    label: app.name,
    submenu: [
      { role: 'about' },
      { type: 'separator' },
      { role: 'services' },
      { type: 'separator' },
      { role: 'hide' },
      { role: 'hideothers' },
      { role: 'unhide' },
      { type: 'separator' },
      { role: 'quit' }
    ]
  }] : []),
  // { role: 'fileMenu' }
  {
    label: 'Application',
    submenu: [
      {
        label: 'About Microsoft Office - Electron',
          click: () =>
                       openAboutWindow({
                           icon_path: 'https://github.com/agam778/Microsoft-Office-Electron/blob/main/icon2.png?raw=true',
                           product_name: 'Microsoft Office - Electron',
                           copyright: 'Copyright (c) 2021 Agampreet Singh Bajaj',
                           package_json_dir: __dirname,
                           bug_report_url: 'https://github.com/agam778/Microsoft-Office-Electron/issues/',
                           bug_link_text: 'Report an issue',
                           adjust_window_size: '2',
                           show_close_button: 'Close',

                     }),
      },
      {
        label: 'Learn More',
        click: async () => {
  const { shell } = require('electron');
  await shell.openExternal('https://github.com/agam778/Microsoft-Office-Electron');
}
},
      {type:'separator'},
      {
        role: 'quit',
        accelerator: process.platform === 'darwin' ? 'Ctrl+Q' : 'Ctrl+Q',
        }

    ]
  },
  // { role: 'editMenu' }
  {
    label: 'Edit',
    submenu: [
      { role: 'undo' },
      { role: 'redo' },
      { type: 'separator' },
      { role: 'cut' },
      { role: 'copy' },
      { role: 'paste' },
      ...(isMac ? [
        { role: 'pasteAndMatchStyle' },
        { role: 'delete' },
        { role: 'selectAll' },
        { type: 'separator' },
        {
          label: 'Speech',
          submenu: [
            { role: 'startSpeaking' },
            { role: 'stopSpeaking' }
          ]
        }
      ] : [
        { role: 'delete' },
        { type: 'separator' },
        { role: 'selectAll' }
      ])
    ]
  },
  // { role: 'viewMenu' }
  {
    label: 'View',
    submenu: [
      { role: 'reload' },
      { role: 'forceReload' },
      { role: 'toggleDevTools' },
      { type: 'separator' },
      { role: 'resetZoom' },
      { role: 'zoomIn',
      accelerator: process.platform === 'darwin' ? 'Control+=' : 'Control+=' },
      { role: 'zoomOut' },
      { type: 'separator' },
      { role: 'togglefullscreen' }
    ]
  },
  // { role: 'windowMenu' }
  {
    label: 'Window',
    submenu: [
      { role: 'minimize' },
      { role: 'zoom' },
      ...(isMac ? [
        { type: 'separator' },
        { role: 'front' },
        { type: 'separator' },
        { role: 'window' }
      ] : [
        { role: 'close' }
      ])
    ]
  }
]

const menu = Menu.buildFromTemplate(template)
Menu.setApplicationMenu(menu)

function createWindow () {
  const win = new BrowserWindow({
    width: 1181,
    height: 670,
    icon: path.join(__dirname, '/icons/icon.png'),
    webPreferences: {
      nodeIntegration: true
    }
  })

  win.loadURL('https://agam778.github.io/Microsoft-Office-Electron/',
{userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.72 Safari/537.36'});
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

try {
  require('electron-reloader')(module)
} catch (_) {}

app.on('ready', function()  {
  autoUpdater.checkForUpdatesAndNotify();
});
