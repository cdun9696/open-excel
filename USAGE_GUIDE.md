# OpenExcel 使用指南

OpenExcel 是一个开源的 Microsoft Excel 加载项，允许你在 Excel 中直接与 LLM（大语言模型）进行对话，支持 OpenAI、Anthropic、Google 等多家服务商，使用你自己的 API 密钥 (BYOK)。

---

## 📦 普通用户安装

如果你只想使用这个插件，无需进行任何开发环境配置。

### 下载清单文件

下载项目中的 [`manifest.prod.xml`](./manifest.prod.xml) 文件到本地。

---

### Windows 平台安装

1. 打开 **Excel**
2. 点击菜单栏 **插入 (Insert)** → **加载项 (Add-ins)** → **我的加载项 (My Add-ins)**
3. 点击 **上传我的加载项 (Upload My Add-in)**
4. 浏览并选择下载的 `manifest.prod.xml` 文件，点击确定
5. 在 **开始 (Home)** 选项卡的功能区中点击 **"Open AI Chat"** 按钮启动插件

---

### macOS 平台安装

1. **复制清单文件到指定目录**
   
   打开终端，运行以下命令：
   ```bash
   # 创建目标目录（如果不存在）
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   
   # 复制清单文件
   cp /path/to/manifest.prod.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```
   
   > ⚠️ 请将 `/path/to/manifest.prod.xml` 替换为你下载文件的实际路径

2. **完全退出并重新打开 Excel**
   
   > 必须完全退出 Excel（Command + Q），不是只关闭窗口

3. 点击菜单栏 **插入 (Insert)** → **加载项 (Add-ins)** → **我的加载项 (My Add-ins)**

4. 在 **共享文件夹 (Shared Folder)** 下选择 **"OpenExcel"**

---

### Excel 网页版安装

1. 打开 [excel.office.com](https://excel.office.com) 并打开一个工作簿
2. 点击 **插入 (Insert)** → **加载项 (Add-ins)** → **更多加载项 (More Add-ins)**
3. 点击 **上传我的加载项 (Upload My Add-in)**
4. 上传 `manifest.prod.xml` 文件

> **注意：** 某些组织的 IT 管理员可能会禁用"上传我的加载项"选项。

---

## 🛠️ 开发者本地开发

如果你想参与开发或进行本地调试，请按照以下步骤操作。

### 环境要求

| 要求 | 说明 |
|------|------|
| **Node.js** | v18 或更高版本 |
| **包管理器** | pnpm（推荐）、npm 或 yarn |
| **Excel** | Microsoft Excel 桌面版 |

### 安装步骤

#### 1. 克隆项目
```bash
git clone https://github.com/huangdaxianer/open-excel.git
cd open-excel
```

#### 2. 安装依赖
```bash
# 使用 pnpm（推荐）
pnpm install

# 或使用 npm
npm install

# 或使用 yarn
yarn install
```

#### 3. 启动开发服务器

**方式一：自动启动 Excel 并加载插件**
```bash
pnpm start
```
此命令会：
- 启动开发服务器（https://localhost:3000）
- 自动打开 Excel
- 自动加载插件到 Excel 中

**方式二：仅启动开发服务器**
```bash
pnpm dev-server
```
此命令仅启动开发服务器，你需要手动将 `manifest.xml` 加载到 Excel 中。

#### 4. 手动加载开发版本的清单文件

如果使用方式二，需要手动加载 `manifest.xml`（注意是 `manifest.xml`，不是 `manifest.prod.xml`）：

**macOS：**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```
然后重启 Excel 并在加载项中选择 "OpenExcel (New)"。

**Windows：**
通过 Excel 的 **插入** → **加载项** → **我的加载项** → **上传我的加载项** 上传 `manifest.xml`。

#### 5. 停止开发
```bash
pnpm stop
```

---

## 📋 常用命令一览

| 命令 | 说明 |
|------|------|
| `pnpm install` | 安装项目依赖 |
| `pnpm start` | 启动开发服务器并自动加载插件到 Excel |
| `pnpm stop` | 停止开发服务器和卸载插件 |
| `pnpm dev-server` | 仅启动开发服务器 (https://localhost:3000) |
| `pnpm build` | 生产环境构建 |
| `pnpm deploy` | 构建并部署到 Cloudflare Pages |
| `pnpm lint` | 运行代码检查 |
| `pnpm lint:fix` | 自动修复代码问题 |
| `pnpm typecheck` | TypeScript 类型检查 |
| `pnpm validate` | 验证 Office 清单文件 |

---

## ⚙️ 配置插件

首次使用时，需要在插件的 **Settings（设置）** 选项卡中进行配置：

### 基本配置

1. **Provider（服务商）** - 选择你的 LLM 服务商
   - OpenAI
   - Anthropic
   - Google
   - 火山引擎 Ark（volcengine-ark）
   - doubao-test-log（内部测试）

2. **API Key（API 密钥）** - 输入对应服务商的 API 密钥

3. **Model（模型）** - 选择要使用的具体模型

> 💡 **提示：** 所有设置都存储在浏览器的 localStorage 中，不会上传到任何服务器。

---

## 🔧 自定义模型配置

如果需要添加自定义模型，可以编辑 `src/taskpane/components/chat/custom-models.ts` 文件：

```typescript
export const CUSTOM_MODELS: Record<string, Model<any>[]> = {
  "你的服务商名称": [
    {
      id: "模型ID",
      name: "模型显示名称",
      provider: "服务商名称",
      api: "anthropic-messages",  // API 格式
      baseUrl: "https://your-api-endpoint.com",
      contextWindow: 64000,
      maxTokens: 8000,
      input: ["text", "image"],
      reasoning: false,
      cost: {
        input: 0.55,
        output: 2.19,
        cacheRead: 0.14,
        cacheWrite: 0.55,
      },
    },
  ],
};
```

修改后需要重新构建或重启开发服务器。

---

## 🛠️ 可用的 Excel 工具

插件提供了以下工具供 AI 操作 Excel：

### 电子表格操作工具（11个）

| 工具 | 说明 |
|------|------|
| `get_cell_ranges` | 读取单元格的值、公式和格式 |
| `get_range_as_csv` | 将数据导出为 CSV 格式（适合分析） |
| `search_data` | 在电子表格中搜索文本 |
| `get_all_objects` | 列出所有图表、数据透视表等对象 |
| `set_cell_range` | 写入值、公式和格式 |
| `clear_cell_range` | 清除单元格内容或格式 |
| `copy_to` | 复制范围（支持公式翻译） |
| `modify_sheet_structure` | 插入/删除/隐藏/冻结行列 |
| `modify_workbook_structure` | 创建/删除/重命名工作表 |
| `resize_range` | 调整列宽和行高 |
| `modify_object` | 创建/更新/删除图表和数据透视表 |

### 扩展工具（1个）

| 工具 | 说明 |
|------|------|
| `eval_officejs` | 在 Excel.run 上下文中执行任意 Office.js 代码 |

---

## 📁 项目结构

```
excel-extension/
├── manifest.xml          # 开发环境清单文件 (localhost:3000)
├── manifest.prod.xml     # 生产环境清单文件 (openexcel.pages.dev)
├── package.json          # 项目配置和脚本
├── webpack.config.js     # Webpack 构建配置
├── tsconfig.json         # TypeScript 配置
├── src/
│   ├── taskpane/         # 任务面板（主要 UI）
│   │   ├── taskpane.html
│   │   ├── index.tsx
│   │   ├── index.css
│   │   └── components/   # React 组件
│   │       └── chat/     # 聊天相关组件
│   ├── commands/         # Office 命令
│   ├── lib/              # 工具库和工具实现
│   └── shims/            # 垫片代码
└── assets/               # 图标资源
```

---

## ❓ 常见问题

### Q: 插件无法加载？
- **macOS：** 确保 `manifest.xml` 或 `manifest.prod.xml` 已复制到正确目录，并完全重启 Excel
- **Windows：** 尝试清除 Office 缓存并重新上传清单文件
- 检查开发服务器是否正在运行（开发模式）

### Q: API 调用失败？
- 检查 API 密钥是否正确
- 确认选择的服务商和模型匹配
- 检查网络连接和代理设置

### Q: 如何切换开发/生产环境？
- 开发环境使用 `manifest.xml`（指向 localhost:3000）
- 生产环境使用 `manifest.prod.xml`（指向 openexcel.pages.dev）

---

## 📄 许可证

MIT

---

## 🔗 相关链接

- [GitHub 仓库](https://github.com/huangdaxianer/open-excel)
- [Office Add-ins 官方文档](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API 文档](https://docs.microsoft.com/en-us/javascript/api/excel)
