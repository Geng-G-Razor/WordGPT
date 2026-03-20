# AGENTS.md

本文件用于帮助后续协作者或代码代理快速接手这个仓库。目标是：先理解这是一个什么项目，再按这个项目自己的运行方式、发布方式和约束来改代码。

## 项目概览

- 这是一个 `Microsoft Word` 的 `Office Add-in`，不是普通网页应用。
- 宿主环境是 `Word + Office.js`，主要界面是右侧 `task pane`。
- 当前前端技术栈是 `React 17 + TypeScript + Fluent UI + webpack`。
- 本仓库虽然作者最初描述为 “WordGPT”，但当前实现已经接入了通过 `/api` 代理访问的模型接口，默认使用 SiliconFlow 兼容接口。

## 先看哪里

开始动手前，优先阅读这些文件：

- `package.json`
- `manifest.xml`
- `webpack.config.js`
- `src/taskpane/index.tsx`
- `src/taskpane/components/App.tsx`
- `src/commands/commands.ts`

## 目录说明

- `src/taskpane/`: Word 侧边栏主界面
- `src/taskpane/components/App.tsx`: 核心业务逻辑，包含密钥保存、模型切换、提示词提交、流式响应处理
- `src/commands/`: Ribbon 命令入口
- `src/stream-debug/`: 流式输出调试页
- `manifest.xml`: Office 插件清单，决定加载入口、权限、按钮文案、资源地址
- `webpack.config.js`: 本地开发与生产构建配置，包含 `/api -> https://api.siliconflow.cn/v1` 代理

## 本地开发

常用命令：

- `pnpm install`
- `pnpm run dev-server`
- `pnpm run start`
- `pnpm run start:desktop`
- `pnpm run stop`
- `pnpm run build`
- `pnpm run validate`
- `pnpm run lint`

说明：

- 开发服务默认跑在 `https://localhost:3000`
- `manifest.xml` 和 webpack 配置都依赖这个地址，本地调试不要随意改端口
- Office Add-in 调试通常依赖开发证书和 sideload 流程，遇到加载失败先检查证书、`manifest.xml`、以及 Word 是否真的加载了最新清单

## 改代码时的核心约束

### 1. 这是 Office Add-in，不是普通单页应用

- 所有 UI 都运行在 Office 宿主内嵌环境里
- 兼容性要保守一些，`tsconfig.json` 当前目标是 `ES5`
- 不要轻易引入对现代浏览器能力要求很高、但未做降级的库

### 2. `manifest.xml` 是高敏感文件

- 改 `taskpane` 地址、图标、按钮文案、权限、命令入口时，必须同步检查 `manifest.xml`
- 生产构建依赖 `webpack.config.js` 中把 `https://localhost:3000/` 替换为 `urlProd`
- 如果新增页面或入口，通常需要同时改 webpack entry、HTML 模板、以及 manifest 资源引用

### 3. Office.js 代码要尊重宿主模型

- 涉及 `Word.run(...)` 时，要明确对象加载、读写范围和 `context.sync()`
- 避免在一次交互里做过多无意义 sync
- 对文档内容的写入属于高风险操作，优先保证“可预期、可撤销、不要误覆盖用户内容”

### 4. 网络请求走本地代理

- 当前前端调用的是相对路径 `/api/chat/completions`
- 真实转发在 `webpack.config.js` 的 `devServer.proxy`
- 如果你修改接口路径、鉴权头或响应结构，要同时检查本地开发和生产部署是否仍然成立

### 5. 本地存储已有约定

- API Key 存在 `localStorage`，键名为 `siliconflowApiKey`
- 模型选择存储键名为 `siliconflowModel`
- 改动这些键名时，要考虑兼容旧数据，不要无提示让现有用户配置失效

## 代码风格建议

- 优先延续现有写法，不要无必要大改框架或目录结构
- 组件改动尽量聚焦，避免把 UI 重构和业务逻辑修改混在一起
- 保持用户可见文案统一，当前主要面向中文用户时，新增提示尽量用中文
- 若新增常量、存储键、模型配置，优先集中声明，避免散落
- 若修改流式输出逻辑，先保护错误处理和降级路径，再做体验优化

## 已知实现特点

- `src/taskpane/components/App.tsx` 是最重要的文件，当前承担了较多状态和请求逻辑
- 项目启用了流式响应解析，包含 SSE `data:` 行处理
- `src/commands/commands.ts` 仍带有较强的模板痕迹，如邮箱通知相关代码；如果要增强 Ribbon 命令，先确认是否真的适用于 Word 宿主
- 仓库里同时存在 `pnpm-lock.yaml` 和 `package-lock.json`，默认优先使用 `pnpm`

## 修改前后的检查清单

完成改动后，至少做这些检查：

- `pnpm run lint`
- `pnpm run build`
- `pnpm run validate`

如果改动涉及以下内容，还应补充验证：

- 文档读写：在 Word 中实际 sideload 后验证
- 流式输出：检查分片、结束标记、错误态、空响应
- manifest：确认按钮能打开 task pane，静态资源地址正确
- 代理请求：确认 `/api` 能正常转发，鉴权头未丢失

## 推荐工作方式

1. 先确认改动属于哪一层：`UI`、`Office 文档交互`、`manifest`、`构建配置`、`模型请求`
2. 只在必要范围内改文件
3. 改完先跑静态检查，再做本地构建
4. 如果涉及 Word 文档写入或插件加载流程，尽量进行一次真实宿主验证

## 不建议直接做的事

- 不要把这个项目当成纯 Web 项目来升级一大批依赖
- 不要在没同步检查 `manifest.xml` 的情况下调整入口页面
- 不要随意修改本地存储键名、代理前缀或本地开发端口
- 不要默认 `commands.ts` 里的模板代码就是正确的 Word 实现

## 给后续 Agent 的一句话

先把自己当成在维护一个“运行在 Word 宿主里的 React 应用”，再去改 UI、请求和文档交互；凡是看起来像普通前端问题的地方，都先确认它会不会被 `Office.js`、`manifest.xml` 或宿主环境约束住。
