# Office Preview for Obsidian

> **🧪 This plugin is currently in alpha testing — issues and feedback are welcome!**
>
> **🧪 本插件正在测试阶段，欢迎提交 Issue 反馈问题和建议！**

[English](#english) | [中文](#chinese)

---

## English {#english}

Preview Word (.docx), PowerPoint (.pptx), and Excel (.xlsx) files directly inside Obsidian — no external applications needed. Supports zoom (30%–300%), one-click text copy, multi-sheet Excel navigation, auto file routing, and adaptive light/dark themes.

### Features

#### Completed

- [x] **DOCX Preview** — paragraphs, headings, tables, images, headers & footers
- [x] **PPTX Preview** — continuous slide scroll, shapes, images, charts
- [x] **XLSX Preview** — multi-sheet tabs, row numbers, frozen headers, data stats
- [x] **Auto Routing** — click Office files in the file tree to open preview
- [x] **Zoom** — toolbar buttons + Ctrl+Scroll, 30%–300%
- [x] **Copy Text** — one-click copy of all preview content
- [x] **Theme Adaptive** — follows Obsidian light/dark theme
- [x] **Mobile Responsive** — works on iOS & Android
- [x] **Context Menu** — right-click "Preview this file"
- [x] **Command Palette** — `Preview current file` command

#### In Development

- [ ] **Real-time Editing** — modify docx/pptx/xlsx content and save directly in Obsidian
  - Text editing
  - Cell data modification
  - Slide content adjustment
- [ ] **Font Rendering Enhancement** — embedded font support, system font fallback
  - CJK font optimization
  - Embedded font parsing
  - Font fallback strategy
- [ ] **Full-text Search Integration** — extract office text into Obsidian search index
- [ ] **Markdown Embed** — embed slides or table ranges via `![[file.pptx#slide3]]`
- [ ] **Legacy Format Support** — .doc / .xls / .ppt (97-2003)
- [ ] **Comments & Track Changes** — display Word comments and revisions
- [ ] **Keyword Highlight** — Ctrl+F search and highlight within preview
- [ ] **Excel Chart Rendering** — render embedded charts in Excel files

### Installation

#### Manual Install

1. Download the latest `main.js`, `styles.css`, `manifest.json`
2. Place them in `.obsidian/plugins/office-preview/` inside your vault
3. Restart Obsidian, enable "Office Preview" in Settings → Community Plugins

```
.vault/
└── .obsidian/
    └── plugins/
        └── office-preview/
            ├── main.js
            ├── styles.css
            └── manifest.json
```

#### Build

```bash
npm install
npm run build
```

### Tech Stack

| File Type | Render Engine |
|-----------|---------------|
| `.docx` | [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) |
| `.pptx` | [pptx-preview](https://www.npmjs.com/package/pptx-preview) |
| `.xlsx` | [SheetJS (xlsx)](https://sheetjs.com/) |

| | |
|---|---|
| Framework | Obsidian Plugin API |
| Language | TypeScript |
| Bundler | esbuild |
| Zoom | CSS `zoom` property |

### Compatibility

| Platform | Minimum Version |
|----------|----------------|
| Obsidian Desktop | 0.15.0+ |
| Obsidian Mobile | 0.15.0+ |
| Windows / macOS / Linux | ✅ |
| iOS / Android | ✅ |

### Shortcuts

| Action | Shortcut |
|--------|----------|
| Zoom In | `Ctrl + Scroll Up` |
| Zoom Out | `Ctrl + Scroll Down` |
| Reset Zoom | Click `1:1` button |
| Preview Current File | Command Palette `Preview current file` |

### Project Structure

```
obsidian-office-preview/
├── main.ts              # Plugin entry + preview view + toolbar + renderers
├── styles.css           # Global styles (DOCX / XLSX / PPTX / Mobile)
├── manifest.json        # Obsidian plugin manifest
├── package.json         # Dependencies
├── esbuild.config.mjs   # Build config
├── tsconfig.json        # TypeScript config
├── versions.json        # Version compatibility table
└── README.md
```

---

## 中文 {#chinese}

在 Obsidian 中直接预览 Word (.docx)、PowerPoint (.pptx) 和 Excel (.xlsx) 文件，无需切换到外部应用。支持缩放（30%–300%）、一键复制文本、Excel 多 Sheet 切换、文件自动路由、明暗主题适配。

### 功能

#### 已完成

- [x] **DOCX 预览** — 段落、标题、表格、图片、页眉页脚
- [x] **PPTX 预览** — 幻灯片连续滚动、形状/图片/图表渲染
- [x] **XLSX 预览** — 多 Sheet 切换、行号、冻结表头、数据统计
- [x] **文件自动路由** — 点击文件树中的 Office 文件自动打开预览
- [x] **缩放** — 工具栏按钮 + Ctrl+滚轮，30%–300%
- [x] **文本复制** — 一键复制预览中的全部文本
- [x] **主题适配** — 跟随 Obsidian 明暗主题自动切换
- [x] **移动端适配** — iOS / Android 均可使用
- [x] **右键菜单** — 文件右键"预览此文件"
- [x] **命令面板** — `预览当前文件` 命令

#### 开发中

- [ ] **实时编辑** — 直接在 Obsidian 中修改 docx/pptx/xlsx 内容并保存
  - 文本编辑
  - 单元格数据修改
  - 幻灯片内容调整
- [ ] **字体渲染增强** — 支持文档内嵌字体、系统字体 fallback
  - 中文字体优化
  - 嵌入字体解析
  - 字体回退策略
- [ ] **全文搜索集成** — 提取 Office 文件文本内容，纳入 Obsidian 搜索索引
- [ ] **Markdown 嵌入** — 通过 `![[file.pptx#slide3]]` 语法嵌入指定幻灯片或表格区域
- [ ] **旧格式支持** — .doc / .xls / .ppt (97-2003)
- [ ] **批注/修订显示** — 显示 Word 文档的批注和修订痕迹
- [ ] **关键字高亮** — 在预览中 Ctrl+F 搜索并高亮
- [ ] **Excel 图表渲染** — 渲染 Excel 文件中的内嵌图表

### 安装

#### 手动安装

1. 下载最新的 `main.js`、`styles.css`、`manifest.json`
2. 放入 Vault 的 `.obsidian/plugins/office-preview/` 目录
3. 重启 Obsidian，在 设置 → 社区插件 中启用 "Office Preview"

```
.vault/
└── .obsidian/
    └── plugins/
        └── office-preview/
            ├── main.js
            ├── styles.css
            └── manifest.json
```

#### 构建

```bash
npm install
npm run build
```

### 技术栈

| 文件类型 | 渲染引擎 |
|---------|----------|
| `.docx` | [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) |
| `.pptx` | [pptx-preview](https://www.npmjs.com/package/pptx-preview) |
| `.xlsx` | [SheetJS (xlsx)](https://sheetjs.com/) |

| | |
|---|---|
| 框架 | Obsidian Plugin API |
| 语言 | TypeScript |
| 构建 | esbuild |
| 缩放 | CSS `zoom` property |

### 兼容性

| 平台 | 最低版本 |
|------|---------|
| Obsidian Desktop | 0.15.0+ |
| Obsidian Mobile | 0.15.0+ |
| Windows / macOS / Linux | ✅ |
| iOS / Android | ✅ |

### 快捷键

| 操作 | 快捷键 |
|------|--------|
| 放大 | `Ctrl + 滚轮上` |
| 缩小 | `Ctrl + 滚轮下` |
| 重置缩放 | 点击 `1:1` 按钮 |
| 预览当前文件 | 命令面板 `预览当前文件` |

### 项目结构

```
obsidian-office-preview/
├── main.ts              # 插件入口 + 预览视图 + 工具栏 + 渲染逻辑
├── styles.css           # 全局样式 (DOCX / XLSX / PPTX / 移动端)
├── manifest.json        # Obsidian 插件清单
├── package.json         # 依赖管理
├── esbuild.config.mjs   # 构建配置
├── tsconfig.json        # TypeScript 配置
├── versions.json        # 版本兼容表
└── README.md
```

---

## License

MIT
