/**
 * scripts/build-manifest.js
 * 根据 deploy.config.js 生成生产环境的 manifest.xml
 * 将 localhost:3000 替换为实际服务器地址
 */
const fs = require("fs");
const path = require("path");

const config = require("../deploy.config.js");
const serverUrl = config.SERVER_URL.replace(/\/+$/, "");

if (!serverUrl || serverUrl.includes("your-server")) {
  console.error("❌ 请先在 deploy.config.js 中设置你的服务器地址！");
  process.exit(1);
}

if (!serverUrl.startsWith("https://")) {
  console.error("❌ 服务器地址必须是 HTTPS！");
  process.exit(1);
}

// 读取开发版 manifest
const manifestPath = path.resolve(__dirname, "../manifest.xml");
let manifest = fs.readFileSync(manifestPath, "utf-8");

// 替换所有 localhost 地址
manifest = manifest.replace(/https:\/\/localhost:3000/g, serverUrl);

// 输出到 dist 目录
const distDir = path.resolve(__dirname, "../dist");
if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir, { recursive: true });
}

const outputPath = path.join(distDir, "manifest.xml");
fs.writeFileSync(outputPath, manifest, "utf-8");

// 生成 install.bat（纯 ASCII，避免 CMD 编码问题）
const installBat = `@echo off
echo.
echo  WordAI Add-in Installer
echo  ========================
echo.
set "MANIFEST_PATH=%~dp0manifest.xml"
if not exist "%MANIFEST_PATH%" (
    echo  [ERROR] manifest.xml not found!
    echo  Please make sure manifest.xml is in the same directory.
    echo.
    pause
    exit /b 1
)
reg add "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "a1b2c3d4-e5f6-7890-abcd-ef1234567890" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1
if %errorlevel% equ 0 (
    echo  [OK] Installation successful!
    echo.
    echo  How to use:
    echo    1. Open Word ^(restart if already open^)
    echo    2. Find the WordAI button in the Home tab
    echo    3. Click to open the sidebar
    echo.
    echo  To uninstall, run uninstall.bat
) else (
    echo  [FAIL] Installation failed. Please run as Administrator.
)
echo.
pause
`;
fs.writeFileSync(path.join(distDir, "install.bat"), installBat);

// 生成 uninstall.bat
const uninstallBat = `@echo off
echo.
echo  WordAI Add-in Uninstaller
echo  ==========================
echo.
reg delete "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "a1b2c3d4-e5f6-7890-abcd-ef1234567890" /f >nul 2>&1
if %errorlevel% equ 0 (
    echo  [OK] Uninstalled successfully! Please restart Word.
) else (
    echo  [INFO] Add-in was not installed or already removed.
)
echo.
pause
`;
fs.writeFileSync(path.join(distDir, "uninstall.bat"), uninstallBat);

console.log(`✅ 生产 manifest 已生成: dist/manifest.xml`);
console.log(`   服务器地址: ${serverUrl}`);
console.log(`✅ install.bat / uninstall.bat 已生成`);
