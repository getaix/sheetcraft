# 概览

sheetcraft 是一个面向导出与模板渲染场景的 Python Excel 库，支持 `.xlsx` 与旧版 `.xls`，提供可定制样式、图片插入、数据验证、公式写入与评估、多工作表操作，以及面向大数据集的性能优化。

- `.xlsx` 默认基于 openpyxl；可选使用 xlsxwriter 提升写入性能（`fast=True`）
- `.xls` 基于 xlwt（有限样式、无数据验证创建能力）
- 模板渲染基于 Jinja2，占位符与 `{% for %}` 重复块
- 可选“格式修复模块”，在保存后对 `.xlsx` 进行结构性修补（不更改数据）

更多快速示例可参考仓库中的 `examples/` 与顶层 `README.md`。

### 预览工具
- 示例预览应用位于 `examples/preview`（Vue3 + Vite + TypeScript）。
- 支持本地选择或拖拽 `.xlsx` 文件进行预览，并在前端轻量修复图片关系路径以提升兼容性。
- 运行方式：进入 `examples/preview` 执行 `npm install && npm run dev`，打开浏览器访问本地地址后选择文件即可。