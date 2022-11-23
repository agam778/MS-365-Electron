<p align="center"><img src="https://github.com/agam778/MS-365-Electron/blob/main/Intro_Image.png?raw=true" alt="Intro Image"></p>

<p align="center">Unofficial Microsoft 365 Web Desktop Wrapper made with Electron</p>

<p align="center">
<a href="https://youtube.com/AgamsTechTricks">
 <img align="center" src="https://img.shields.io/badge/Made%20With%20♥-by%20Agam-orange?style=style=flat">   
 </a>
<a href="https://electronjs.org">
 <img align="center" src="https://img.shields.io/badge/Developed%20With-Electron-red?logo=Electron&logoColor=white&style=flat">  
 </a>
<a href="https://github.com/agam778/MS-365-Electron/blob/main/LICENSE">
 <img align="center" src="https://img.shields.io/github/license/agam778/MS-365-Electron?style=flat">  
 </a>
<a  href="https://github.com/agam778/MS-365-Electron/releases/">
 <img align="center" src="https://img.shields.io/github/v/release/agam778/MS-365-Electron?label=Release&logo=github&style=style=flat&color=blue">  
 </a>
<a href="https://github.com/agam778/MS-365-Electron/releases/">
 <img align="center" src="https://img.shields.io/github/downloads/agam778/MS-365-Electron/total?label=Downloads&style=style=flat">
 </a>
 <a href="https://github.com/agam778/MS-365-Electron/releases/latest/">
 <img align="center" src="https://img.shields.io/github/downloads/agam778/MS-365-Electron/latest/total?label=Downloads%40Latest">
 </a>
 <a href="https://github.com/agam778/MS-365-Electron/actions/workflows/build.yml">
  <img align="center" src="https://github.com/agam778/MS-365-Electron/actions/workflows/build.yml/badge.svg">
 </a>
</p>

# Table of contents

- [Table of contents](#table-of-contents)
- [Introduction](#introduction)
- [Windows](#windows)
  - [💿 Windows Installation](#-windows-installation)
  - [📸 Windows Preview](#-windows-preview)
- [macOS](#macos)
  - [💿 macOS Installation](#-macos-installation)
- [Linux](#linux)
  - [💿 Linux Installation](#-linux-installation)
    - [Ubuntu/Debian based distribution installation](#ubuntudebian-based-distribution-installation)
    - [Red Hat/Fedora based distribution installation](#red-hatfedora-based-distribution-installation)
    - [Arch/Manjaro Linux based distribution installation](#archmanjaro-linux-based-distribution-installation)
  - [📸 Ubuntu Preview](#-ubuntu-preview)
- [💻 Developing Locally](#-developing-locally)
- [📃 MIT License](#-mit-license)
      - [*Disclaimer: Not affiliated with Microsoft. Office, the name, website, images/icons are the intellectual properties of Microsoft.*](#disclaimer-not-affiliated-with-microsoft-office-the-name-website-imagesicons-are-the-intellectual-properties-of-microsoft)

# Introduction

This project is basically a Desktop wrapper for the web version of [Microsoft 365](https://microsoft365.com), which is free but with some basic limits.

I initially made this project because I wanted to use Microsoft 365 on my Linux machine, so to get the feel of using a native app while using it on Linux, I thought to made this project. Later, I decided to make it public so that others can enjoy this too!

Don't expect this to be a full-fledged Microsoft 365 Desktop Suite (like we have for Windows/macOS), it's just a wrapper of the web version of Microsoft 365.

Note - Windows Hello or Sign in with Security key is **not** supported and will show you an error. You will have to manually sign in with your E-Mail and Password.

***Do Expect bugs***

Supported Platforms

1. Windows x64 (EXE File)
2. macOS x64 (DMG File)
3. Ubuntu/Debian based distributions (DEB File)
4. Red Hat Linux/Fedora based distributions (RPM File)
5. Arch/Manjaro Linux based distributions (Uploaded on AUR)
6. Gentoo Linux (Unofficial overlay)
7. All Distributions supporting AppImage (AppImage File); and
8. All Distributions supporting Snap (Uploaded on Snap Store)
9. 
# Windows

## 💿 Windows Installation

For Installing this app on Windows :- 

1) Just go to the [Releases](https://github.com/agam778/MS-365-Electron/releases) page
2) Scroll down and click the  `MS-365-Electron-vx.x.x-win-x64.exe` file. The Setup file will start downloading.
3) After it downloads, click on the file and proceed with the Installation. You can choose whether to install for only you or all the users on the PC. You can always start the app from Start Menu or from the Desktop Shortcut.

## 📸 Windows Preview

[Click Here](https://github.com/agam778/MS-365-Electron/blob/main/Preview/Windows%20Preview.png?raw=true)

# macOS

## 💿 macOS Installation

For Installing this app on Mac :-

1. Just go to the [Releases](https://github.com/agam778/MS-365-Electron/releases) page
2. Scroll down and click the `.dmg` file. The build is only for Intel Macs.
3. After it downloads, click on the file and mount it on your system. Now drag my app to the Applications Folder (There will be a shortcut in the opened window too) and your app will be installed. Open from Launchpad and enjoy.

# Linux

## 💿 Linux Installation

<a href="https://snapcraft.io/ms-office-electron">
  <img alt="Get it from the Snap Store" src="https://snapcraft.io/static/images/badges/en/snap-store-black.svg" />
</a>

### Ubuntu/Debian based distribution installation

For Installing in Ubuntu/Debian based distribution :- 

1) Go to the [Releases](https://github.com/agam778/MS-365-Electron/releases) page
2) Scroll down and click the `.deb` file to download it.
3) Then run the deb file and click Install to install the App. Launch it from the Applications Menu.

### Red Hat/Fedora based distribution installation

For Installing in Red Hat/Fedora based distribution :- 

1) Go to the [Releases](https://github.com/agam778/MS-365-Electron/releases) page
2) Scroll down and click the `.rpm` file to download it.
3) Then run the rpm file and click Install to install the App. Launch it from the Applications Menu.

### Arch/Manjaro Linux based distribution installation

Arch Linux builds have been published to "AUR" now!

1. Install any AUR helper like [`yay`](https://github.com/Jguer/yay)

2. There are 2 packages in the AUR
   - `ms-office-electron-bin`: For installing pre-built releases
   - `ms-office-electron-git`: For building the app from source and installing.

3. Now, for example, using `yay`, run:
   ```bash
   yay -Sy ms-office-electron-*
   ```
   To install the package accordingly.

4. Wait for it to install and tada! The app is installed.

If you find any issues in using the AUR Builds, please create a [New Issue](https://github.com/agam778/MS-365-Electron/issues/new) and i'll try to fix that as soon as possible :D

## 📸 Ubuntu Preview

[Click Here](https://github.com/agam778/MS-365-Electron/blob/main/Preview/Ubuntu%20Preview.png?raw=true)

# 💻 Developing Locally
To build the app locally:<br>
Run this script to automatically install `nodejs`, `yarn` and all the dependencies, and automatically start/build the app (it will show options) (Note: for Linux and macOS Only!):
```bash
git clone https://github.com/agam778/MS-365-Electron.git
cd MS-365-Electron
bash build.sh
```
<br>
Or:<br>
Run the following commands to clone the repository and install the dependencies

```bash
git clone https://github.com/agam778/MS-365-Electron.git
cd MS-365-Electron
yarn install
```
```bash
$ yarn run
yarn run v1.22.18
info Commands available from binary scripts: asar, dircompare, ejs, electron, electron-builder, electron-osx-flat, electron-osx-sign, extract-zip, install-app-deps, is-ci, jake, js-yaml, json5, mime, mkdirp, node-which, rc, rimraf, semver
info Project commands
   - dist
      electron-builder
   - pack
      electron-builder --dir
   - start
      electron .
question Which command would you like to run?:
```

To start the app, run `yarn start`<br>
To build the app, run `yarn dist`

# 📃 MIT License

#### *Disclaimer: Not affiliated with Microsoft. Office, the name, website, images/icons are the intellectual properties of Microsoft.*
