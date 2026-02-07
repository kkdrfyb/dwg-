# cad-text-replace

## 适用场景
- DWG/DXF 批量文字替换
- 支持 TEXT/MTEXT/ATTRIB/块内文字

## 推荐实现策略
- 使用 ODA 将 DWG 转 DXF
- 结构化解析 + 规则替换
- 生成替换后的 DXF，再通过 ODA 反向转回 DWG

## 任务分解建议
- 替换规则配置与预览
- 批量替换与回滚
- 结果校验与日志输出

