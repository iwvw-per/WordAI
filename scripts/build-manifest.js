/**
 * scripts/build-manifest.js
 * 根据 deploy.config.js 生成生产环境的 manifest.xml
 * 将 localhost:3000 替换为实际服务器地址
 */
const fs = require("fs");
const path = require("path");

// 优先尝试从环境变量读取（适配 Docker 构建），否则读取配置文件
let serverUrl = process.env.SERVER_URL;
if (!serverUrl) {
  try {
    const config = require("../deploy.config.js");
    serverUrl = config.SERVER_URL;
  } catch (e) {
    // 忽略配置文件缺失
  }
}

if (serverUrl) {
  serverUrl = serverUrl.replace(/\/+$/, "");
}

if (!serverUrl || serverUrl.includes("your-server")) {
  console.error("❌ 错误：未发现有效的 SERVER_URL！");
  console.error("   - 如果是本地构建，请检查 deploy.config.js 是否包含 SERVER_URL。");
  console.error("   - 如果是 CI 构建，请确保 GitHub Secrets 中设置了 SERVER_URL，且 Dockerfile 正确传递了该变量。");
  process.exit(1);
}

if (!serverUrl.startsWith("https://") && serverUrl !== "__SERVER_URL__") {
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

echo  1. Registering developer manifest...
reg add "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "a1b2c3d4-e5f6-7890-abcd-ef1234567890" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

if %errorlevel% equ 0 (
    echo  [OK] Installation successful!
    echo.
    echo  If the button disappears after restart:
    echo    1. Go to Word -^> Insert -^> My Add-ins
    echo    2. Look for "WordAI" under "Developer Add-ins"
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

// 生成 clear-cache.bat (解决重启消失或脚本不更新的终极方案)
const clearCacheBat = `@echo off
echo.
echo  WordAI - Clearing Office Cache
echo  ==============================
echo.
echo  Closing Word and Excel...
taskkill /f /im winword.exe >nul 2>&1
taskkill /f /im excel.exe >nul 2>&1
echo  Clearing WEF cache folders...
rmdir /s /q "%LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Wef" >nul 2>&1
rmdir /s /q "%AppData%\\Microsoft\\Office\\16.0\\WEF" >nul 2>&1
echo.
echo  [OK] Cache cleared! Please restart Word.
echo.
pause
`;
fs.writeFileSync(path.join(distDir, "clear-cache.bat"), clearCacheBat);

console.log(`✅ 生产 manifest 已生成: dist/manifest.xml`);
console.log(`   服务器地址: ${serverUrl}`);
console.log(`✅ install.bat / uninstall.bat / clear-cache.bat 已生成`);
