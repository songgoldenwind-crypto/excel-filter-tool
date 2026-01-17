# 在Mac上打包Windows exe文件 - 完整指南

## 步骤1：安装Homebrew（如果已安装请跳过）

打开终端，运行：
```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

## 步骤2：安装Wine

Wine可以在Mac上运行Windows程序。在终端运行：

```bash
brew install --cask wine-stable
```

这个安装过程可能需要10-20分钟，请耐心等待。

## 步骤3：验证Wine安装

安装完成后，运行以下命令验证：
```bash
wine --version
```

如果显示版本号（例如 Wine 7.0 或类似），说明安装成功。

## 步骤4：下载Windows版Python

由于PyInstaller需要特定平台的Python，我们需要下载Windows版本：

1. 访问：https://www.python.org/downloads/windows/
2. 下载 Python 3.10.x 或 3.11.x 的 Windows embeddable package (32-bit或64-bit都可以)
3. 下载后会得到类似 `python-3.10.x-embed-amd64.zip` 的文件

## 步骤5：准备Wine环境

在终端运行以下命令：

```bash
# 创建Wine前缀
export WINEPREFIX=~/.wine_pyinstaller
wineboot -i

# 等待Wine初始化完成
```

## 步骤6：安装Python到Wine环境

假设您下载的Python在 ~/Downloads 目录：

```bash
# 解压Python
cd ~/Downloads
unzip python-3.*-embed-amd64.zip -d python_windows

# 复制��Wine环境
mkdir -p ~/.wine_pyinstaller/drive_c/python
cp -r ~/Downloads/python_windows/* ~/.wine_pyinstaller/drive_c/python/

# 创建Python启动脚本
cat > ~/.wine_pyinstaller/drive_c/python/python.bat << 'EOF'
@echo off
set PYTHONHOME=C:\python
C:\python\python.exe %*
EOF
```

## 步骤7：安装pip和依赖

```bash
# 下载get-pip.py
curl https://bootstrap.pypa.io/get-pip.py -o /tmp/get-pip.py

# 在Wine中运行
wine ~/.wine_pyinstaller/drive_c/python/python.exe /tmp/get-pip.py

# 安装依赖
wine ~/.wine_pyinstaller/drive_c/python/Scripts/pip.exe install openpyxl pyinstaller
```

## 步骤8：打包程序

现在您可以使用提供的 `build_with_wine.sh` 脚本打包了！

```bash
cd /Users/songjinfeng/Desktop/shuju
bash build_with_wine.sh
```

打包完成后，在 `dist` 目录下会生成 `.exe` 文件。

## 常见问题

### Q: Wine安装失败
A: 确保您的Mac OS版本较新（10.13+），如果仍然失败，可能需要使用XQuartz：
```bash
brew install --cask xquartz
```

### Q: 打包过程很慢
A: 正常现象，Wine模拟Windows环境会比较慢，请耐心等待。

### Q: 打包出来的exe在Windows上运行报错
A: 可能是Windows Defender误报，可以添加信任或临时关闭杀毒软件。

## 替代方案：使用GitHub Actions自动打包

如果上述方法太复杂，我可以帮您设置GitHub Actions，可以自动在云端打包Windows exe，无需本地配置。
