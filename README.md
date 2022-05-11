<p align="center"><img src="https://github.com/agam778/Microsoft-Office-Electron/blob/main/Intro_Image.png?raw=true" alt="Intro Image"></p>

<p align="center">An Unofficial Microsoft Office Online Desktop Client made with Electron. Free of Cost.</p>

<p align="center">
<a href="https://bit.ly/agamtechtricks">
 <img align="center" src="https://img.shields.io/badge/Made%20With%20♥-by%20Agam-orange?style=style=flat">   
 </a>
<a href="https://electronjs.org">
 <img align="center" src="https://img.shields.io/badge/Developed%20With-Electron-red?logo=Electron&logoColor=white&style=flat">  
 </a>
<a href="https://github.com/agam778/MS-Office-Electron/blob/main/LICENSE">
 <img align="center" src="https://img.shields.io/github/license/agam778/MS-Office-Electron?style=flat">  
 </a>
<a  href="https://github.com/agam778/MS-Office-Electron/releases/">
 <img align="center" src="https://img.shields.io/github/v/release/agam778/MS-Office-Electron?label=Release&logo=github&style=style=flat&color=blue">  
 </a>
<a href="https://github.com/agam778/MS-Office-Electron/releases/">
 <img align="center" src="https://img.shields.io/github/downloads/agam778/MS-Office-Electron/total?label=Downloads&style=style=flat">
 </a>
 <a href="https://github.com/agam778/MS-Office-Electron/releases/latest/">
 <img align="center" src="https://img.shields.io/github/downloads/agam778/MS-Office-Electron/latest/total?label=Downloads%40Latest">
 </a>
 <a href="https://github.com/agam778/MS-Office-Electron/actions/workflows/build.yml">
  <img align="center" src="https://github.com/agam778/MS-Office-Electron/actions/workflows/build.yml/badge.svg">
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

As we know that Microsoft Office is a paid service, they also have a free Web version [here](https://office.com).

For those people who can't afford Microsoft Office Subscription, or don't want to pay for that, or like the web version of it, Then **MS-Office-Electron** is the app you require.

**MS-Office-Electron** is the app in which the Web Version of Microsoft Office is wrapped into a Cross-Platform App with Electron. And all the Services of Microsoft Office will work for free.

Note - Windows Hello or Sign in with Security key is not currently supported and will show you an error. You will have to manually sign in with your E-Mail and Password.

***Do Expect bugs***

Supported Platforms

1. Windows x64 (EXE File)
2. Ubuntu/Debian based distributions (DEB File)
3. Red Hat Linux/Fedora based distributions (RPM File)
4. Arch/Manjaro Linux based distributions (PKG.TAR.ZST File)
5. All Distributions supporting AppImage (AppImage File)

Arch Linux builds are now on AUR!<br>
Mac OS is supported now! See the instructions how to install below<br>

# Windows

## 💿 Windows Installation

For Installing this app on Windows :- 

1) Just go to the [Releases](https://github.com/agam778/MS-Office-Electron/releases) page
2) Scroll down and click the  `setup.x.x.x.exe` file. The Setup file will start downloading.
3) After it downloads, click on the file and proceed with the Installation. You can choose whether to install for only you or all the users on the PC. You can always start the app from Start Menu or from the Desktop Shortcut.

## 📸 Windows Preview

[Click Here](https://github.com/agam778/MS-Office-Electron/blob/main/Preview/Windows%20Preview.png?raw=true)

# macOS

## 💿 macOS Installation

For Installing this app on Mac :-

1. Just go to the [Releases](https://github.com/agam778/MS-Office-Electron/releases) page
2. Scroll down and click the `.dmg` file. The build is only for Intel Macs.
3. After it downloads, click on the file and mount it on your system. Now drag my app to the Applications Folder (There will be a shortcut in the opened window too) and your app will be installed. Open from Launchpad and enjoy.

# Linux

## 💿 Linux Installation

### Ubuntu/Debian based distribution installation

For Installing in Ubuntu/Debian based distribution :- 

1) Go to the [Releases](https://github.com/agam778/MS-Office-Electron/releases) page
2) Scroll down and click the `.deb` file to download it.
3) Then run the deb file and click Install to install the App. Launch it from the Applications Menu.

### Red Hat/Fedora based distribution installation

For Installing in Red Hat/Fedora based distribution :- 

1) Go to the [Releases](https://github.com/agam778/MS-Office-Electron/releases) page
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

If you find any issues in using the AUR Builds, please create a [New Issue](https://github.com/agam778/MS-Office-Electron/issues/new) and i'll try to fix that as soon as possible :D

## 📸 Ubuntu Preview

[Click Here](https://github.com/agam778/MS-Office-Electron/blob/main/Preview/Ubuntu%20Preview.png?raw=true)

# 💻 Developing Locally
To build the app locally:<br>
Run this script to automatically install `nodejs`, `yarn` and all the dependencies, and automatically start/build the app (it will show options) (Note: for Linux and macOS Only!):
```bash
git clone --depth=1 https://github.com/agam778/MS-Office-Electron.git
cd MS-Office-Electron
bash build.sh
```
<br>
Or:<br>
Run the following commands to clone the repository and install the dependencies

```bash
git clone https://github.com/agam778/MS-0ffice-Electron.git
cd MS-Office-Electron
yarn install
```
```bash
$ yarn run
yarn run v1.22.17
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
