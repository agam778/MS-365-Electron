#!/bin/bash
if ! [ -x "$(command -v node)" ]; then
    echo 'Error: nodejs is not installed.' >&2
    echo 'Installing nodejs now; this may take a while.'
    if [ "$(uname)" == "Linux" ]; then
        if [ "$(id -u)" != "0" ]; then
            if [ -f /etc/debian_version ]; then
                curl -fsSL https://deb.nodesource.com/setup_17.x | sudo -E bash -
                sudo apt-get install -y nodejs
                elif [ -f /etc/redhat-release ]; then
                sudo yum install nodejs
                elif [ -f /etc/arch-release ]; then
                sudo pacman -S nodejs
                elif [ -f /etc/gentoo-release ]; then
                sudo emerge nodejs
                elif [ -f /etc/SuSE-release ]; then
                sudo zypper install nodejs
                elif [ -f /etc/fedora-release ]; then
                sudo dnf install nodejs
                elif [ -f /etc/centos-release ]; then
                sudo yum install nodejs
                elif [ -f /etc/nixos ]; then
                sudo nix-env -iA nodejs
            fi
        else
            if [ -f /etc/debian_version ]; then
                curl -fsSL https://deb.nodesource.com/setup_17.x | bash -
                apt-get install -y nodejs
                elif [ -f /etc/redhat-release ]; then
                yum install nodejs
                elif [ -f /etc/arch-release ]; then
                pacman -S nodejs
                elif [ -f /etc/gentoo-release ]; then
                emerge nodejs
                elif [ -f /etc/SuSE-release ]; then
                zypper install nodejs
                elif [ -f /etc/fedora-release ]; then
                dnf install nodejs
                elif [ -f /etc/centos-release ]; then
                yum install nodejs
                elif [ -f /etc/nixos ]; then
                nix-env -iA nodejs
            fi
        fi
        elif [ "$(uname)" == "Darwin" ]; then
        brew install node
        elif [ "$(uname)" == "MINGW32_NT-10.0" ]; then
        echo 'Error: nodejs is not installed.' >&2
        echo 'Please install nodejs manually.'
        exit 0
    fi
fi

if ! [ -x "$(command -v yarn)" ]; then
    echo 'Error: yarn is not installed.' >&2
    echo 'Installing yarn now; this may take a while.'
    if [ "$(uname)" == "Linux" ]; then
        if [ "$(id -u)" != "0" ]; then
            if [ -f /etc/debian_version ]; then
                sudo apt-get install -y yarn
                elif [ -f /etc/redhat-release ]; then
                sudo yum install yarn
                elif [ -f /etc/arch-release ]; then
                sudo pacman -S yarn
                elif [ -f /etc/gentoo-release ]; then
                sudo emerge yarn
                elif [ -f /etc/SuSE-release ]; then
                sudo zypper install yarn
                elif [ -f /etc/fedora-release ]; then
                sudo dnf install yarn
                elif [ -f /etc/centos-release ]; then
                sudo yum install yarn
                elif [ -f /etc/nixos ]; then
                sudo nix-env -iA yarn
            fi
        else
            if [ -f /etc/debian_version ]; then
                apt-get install -y yarn
                elif [ -f /etc/redhat-release ]; then
                yum install yarn
                elif [ -f /etc/arch-release ]; then
                pacman -S yarn
                elif [ -f /etc/gentoo-release ]; then
                emerge yarn
                elif [ -f /etc/SuSE-release ]; then
                zypper install yarn
                elif [ -f /etc/fedora-release ]; then
                dnf install yarn
                elif [ -f /etc/centos-release ]; then
                yum install yarn
                elif [ -f /etc/nixos ]; then
                nix-env -iA yarn
            fi
        fi
        elif [ "$(uname)" == "Darwin" ]; then
        brew install yarn
        elif [ "$(uname)" == "MINGW32_NT-10.0" ]; then
        echo 'Error: yarn is not installed.' >&2
        echo 'Please install yarn manually.'
        exit 0
    fi
fi

if [ -d "./.git" ]; then
    echo "Detected a cloned repository, Continuing..."
else
    echo "Repository not found, cloning now..."
    if ! [ -x "$(command -v git)" ]; then
        echo 'Error: git is not installed.' >&2
        echo 'Installing git now; this may take a while.'
        if [ "$(uname)" == "Linux" ]; then
            if [ "$(id -u)" == "0" ]; then
                if [ -f /etc/debian_version ]; then
                    apt-get install git
                    elif [ -f /etc/redhat-release ]; then
                    yum install git
                    elif [ -f /etc/arch-release ]; then
                    pacman -S git
                    elif [ -f /etc/gentoo-release ]; then
                    emerge git
                    elif [ -f /etc/SuSE-release ]; then
                    zypper install git
                    elif [ -f /etc/fedora-release ]; then
                    dnf install git
                    elif [ -f /etc/centos-release ]; then
                    yum install git
                    elif [ -f /etc/nixos ]; then
                    nix-env -iA git
                fi
            else
                if [ -f /etc/debian_version ]; then
                    sudo apt-get install git
                    elif [ -f /etc/redhat-release ]; then
                    sudo yum install git
                    elif [ -f /etc/arch-release ]; then
                    sudo pacman -S git
                    elif [ -f /etc/gentoo-release ]; then
                    sudo emerge git
                    elif [ -f /etc/SuSE-release ]; then
                    sudo zypper install git
                    elif [ -f /etc/fedora-release ]; then
                    sudo dnf install git
                    elif [ -f /etc/centos-release ]; then
                    sudo yum install git
                    elif [ -f /etc/nixos ]; then
                    sudo nix-env -iA git
                fi
            fi
            elif [ "$(uname)" == "Darwin" ]; then
            brew install git
            elif [ "$(uname)" == "MINGW32_NT-10.0" ]; then
            echo 'Error: git is not installed.' >&2
            echo 'Please install git manually.'
        fi
    fi
    git clone --depth=1 https://github.com/agam778/MS-Office-Electron; cd MS-Office-Electron
    echo 'Cloned the repository'
fi

clear
echo 'Installing Dependencies'
if [ "$(id -u)" != "0" ]; then
    sudo yarn install
else
    yarn install
fi

clear
echo 'Do you want to run the app or build the app?'
echo '1. Run the app'
echo '2. Build the app'
echo '3. Exit'
echo 'Enter your choice:'; read choice;
if [ $choice -eq 1 ]; then
    echo 'Running the app...'
    yarn start
    elif [ $choice -eq 2 ]; then
    echo 'Building the app...'
    if [ "$(id -u)" != "0" ]; then
        if [ "$(uname -m)" == "arm64" ]; then
            sudo yarn dist --arm64
            elif [ "$(uname -m)" == "x86_64" ]; then
            sudo yarn dist --x64
        fi
    else
        if [ "$(uname -m)" == "arm64" ]; then
            yarn dist --arm64
            elif [ "$(uname -m)" == "x86_64" ]; then
            yarn dist --x64
        fi
    fi
    elif [ $choice -eq 3 ]; then
    echo 'Exiting...'
    exit 1
fi

echo 'Finished successfully! 🎉 '
exit 0
