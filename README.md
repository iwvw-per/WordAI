# WordAI - Word LLM 智能文字处理插件

⚡ 在 Word 中使用自定义 LLM API 一键处理文字。

## 快速开始（开发）

```bash
npm install     # 安装依赖
npm start       # 一键启动（服务器 + Word + 插件加载）
```

## 部署到服务器 (Docker 推荐)

使用 Docker Compose 可以最快地完成部署，并支持运行时通过环境变量修改服务器地址。

### 1. 编写 `docker-compose.yml`

```yaml
services:
  wordai:
    image: ghcr.io/iwvw/wordai:latest
    ports:
      - "8080:80"
    environment:
      - SERVER_URL=${SERVER_URL:-https://your-domain.com}
    restart: always
```

### 2. 启动

```bash
docker compose up -d
```

> [!TIP]
> 启动后，用户直接访问 `http://your-server-ip:8080/manifest.xml` 即可下载已配置好的插件清单。

---

## 部署到服务器 (传统方式)

### 第一步：配置服务器地址

编辑 `deploy.config.js`，填入你的服务器地址：

```js
module.exports = {
  SERVER_URL: "https://your-server.com/wordai",
};
```

### 第二步：构建

```bash
npm run build:prod
```

构建完成后 `dist/` 目录包含：
- `taskpane.html` + `taskpane.bundle.js` + `assets/` → 上传到服务器
- `manifest.xml` → 已自动替换为服务器地址
- `install.bat` / `uninstall.bat` → 给用户的安装脚本

### 第三步：部署

将 `dist/` 中的网页文件上传到你的服务器对应路径。

### 第四步：分发

给用户的文件只需要两个：
```
manifest.xml    ← 插件清单
install.bat     ← 双击安装
uninstall.bat   ← 双击卸载
```

用户在新电脑上：**双击 install.bat → 重启 Word → 完成**。

## 功能

- ⚡ **快捷操作** - 润色、降AI、改写、翻译、纠错、缩写、扩写
- 🔧 **自定义 API** - 任何 OpenAI 兼容格式 of LLM 端点
- 🤖 **自动获取模型** - 从 API 端点自动拉取模型列表
- 📝 **自定义提示词** - 添加/编辑/删除提示词模板
- 🎨 **格式保留** - OOXML 级别操作，保留加粗、斜体等
- ⏭️ **智能跳过** - 自动跳过标题、表格、公式、交叉引用

## 常用命令

| 命令 | 说明 |
|------|------|
| `npm start` | 一键启动开发环境 |
| `npm stop` | 停止开发环境 |
| `npm run dev` | 仅启动开发服务器 |
| `npm run build:prod` | 构建生产版本（含 manifest 替换） |
| `npm run sideload` | 仅注册插件到 Word |
| `npm run unload` | 仅卸载插件 |
