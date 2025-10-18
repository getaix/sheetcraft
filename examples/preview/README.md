# 预览应用（Vue3 + Vite + TypeScript）

该预览应用用于本地查看 `.xlsx` 工作簿的内容，现已重构为 TypeScript 与 Vue3 Composition API，并对 UI 与交互进行优化。

主要特性
- 支持通过“选择文件”或拖拽方式导入本地 `.xlsx` 文件预览。
- 自动修复部分图片关系路径（将 `media/`、`../media/` 等统一为 `/xl/media/`），提升前端解析兼容性。
- 统一中文提示与错误文案，交互更友好。

使用方法
1. 在项目根目录运行：`npm run dev`（或通过上层工具启动）。
2. 打开浏览器访问开发地址，点击“选择文件”或直接拖拽本地 `.xlsx` 文件到页面中。
3. 若预览异常，可尝试在本地用 Excel 打开，确认文件内容与图片关系是否存在问题。

文件结构
- `src/App.vue`：预览主组件（TypeScript + Composition API）。
- `src/main.ts`：应用入口文件。
- `tsconfig.json`：TypeScript 编译配置。

注意事项
- 当前图片关系修复仅做轻量处理，覆盖常见路径前缀缺失场景。如需更复杂的兼容，请在后端或生成侧规范关系文件。
