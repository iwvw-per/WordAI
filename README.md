# WordAI - Word LLM 智能文字处理插件

⚡ 在 Word 中使用自定义 LLM API 一键处理文字。

> [!NOTE]
> 本项目已迁移至组织模式进行维护：[iwvw-per/WordAI](https://github.com/iwvw-per/WordAI)

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
    image: iwvw-per/wordai:latest
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

## 如何强制更新（解决缓存问题）

如果您在服务端更新了代码，但用户端显示的依然是旧版本（如版本号显示的还是 v1.0.0），请尝试以下方案：

### 方案 A：自动更新（推荐）
本项目已启用 **Webpack ContentHash** 技术。每次服务端代码变动都会生成全新的文件名，Word 会在下次打开时自动拉取新版。

### 方案 B：手动清理环境（彻底方案）
如果自动更新未触发，请让用户运行 `dist/` 目录下的 **`clear-cache.bat`**。
该脚本会：
1. 强制关闭 Word 进程。
2. 删除 Office 浏览器的底层缓存文件夹。
3. 再次打开 Word 时，插件将 100% 同步服务器最新状态。

---

## 功能

- ⚡ **快捷操作** - 润色、降AI、改写、翻译、纠错、缩写、扩写
- 🔧 **自定义 API** - 任何 OpenAI 兼容格式 of LLM 端点
- 🤖 **自动获取模型** - 从 API 端点自动拉取模型列表
- 📝 **自定义提示词** - 添加/编辑/删除提示词模板
- 🎨 **格式保留** - OOXML 级别操作，保留加粗、斜体等
- ⏭️ **智能跳过** - 自动跳过标题、表格、公式、交叉引用、图片、目录

## 常用命令

| 命令 | 说明 |
|------|------|
| `npm start` | 一键启动开发环境 |
| `npm stop` | 停止开发环境 |
| `npm run dev` | 仅启动开发服务器 |
| `npm run build:prod` | 构建生产版本（含缓存破坏机制） |
| `npm run sideload` | 仅注册插件到 Word |
| `npm run unload` | 仅卸载插件 |
