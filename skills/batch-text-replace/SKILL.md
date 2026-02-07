# batch-text-replace

## 适用场景
- Word/Excel 批量文字替换
- 支持多关键字与规则替换

## 推荐实现策略
- Word 使用 OpenXML
- Excel 使用 ClosedXML 或 EPPlus
- 提供替换预览与结果日志

## 任务分解建议
- 规则配置与校验
- 批量执行与回滚
- 结果差异报告

## 协作与发布
- 不在本地自动执行 `git push`
- 如需自动发布，使用 GitHub Actions 统一处理
