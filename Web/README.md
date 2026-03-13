# 部署指南

您已生成了一个简单的下载页面 `index.html`。

## 1. 准备文件
请运行 `Prepare-Site.ps1` 脚本。它会自动将网页文件和最新的安装包复制到 `Web\output` 文件夹中，并将安装包重命名为 `QSBar_Setup.exe` 以匹配网页链接。

## 2. 关于 IP 地址
您提供的 `172.24.166.148` 是阿里云服务器的**内网 IP**，外网无法直接访问。
请登录阿里云控制台，找到该实例的 **公网 IP (Public IP)**。通常是一个类似 `47.xxx.xxx.xxx` 或 `120.xxx.xxx.xxx` 的地址。

## 3. 上传与部署

### 如果是 Windows 服务器 (使用远程桌面)
1. 远程连接到服务器 (使用公网 IP)。
2. 将 `Web\output` 文件夹复制到服务器上的某个位置 (例如 `C:\QSBarWeb`)。
3. **方式一 (最简单)**: 如果服务器安装了 Python，在文件夹内打开 CMD，运行 `python -m http.server 80`。
4. **方式二 (推荐)**: 使用 IIS (Internet Information Services) 添加一个新网站，指向该目录。

### 如果是 Linux 服务器 (使用 SSH)
1. 使用 SCP 上传文件:
   `scp -r Web/output/* root@<公网IP>:/var/www/html/`
   (如果是 Nginx/Apache 默认目录)
2. 确保 Nginx 或 Apache 正在运行。
3. 访问 `http://<公网IP>/` 即可看到下载页面。

## 4. 测试
在浏览器中输入 `http://<您的公网IP>/`，应该能看到下载页面。点击“下载”按钮即可下载安装包。
