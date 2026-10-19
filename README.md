# 雪球 · 按股票找持仓组合

Tampermonkey 用户脚本。按股票中文名称检索雪球公开模拟组合，支持完整结果排序、批量查询关停状态、主动调仓记录和当前持仓核验，以及 Excel 导出。

## 安装及自动更新

**[点击安装最新脚本](https://raw.githubusercontent.com/wsvv/xueqiu_stock_cube_finder/main/xueqiu_stock_cube_finder.user.js)**

先安装浏览器扩展 Tampermonkey，再打开上面的链接确认安装。已在本地安装旧版的用户，建议先备份 Tampermonkey 脚本及数据，避免生成重复脚本。

脚本元数据包含 `@updateURL` 和 `@downloadURL`，均指向仓库 `main` 分支的 Raw 文件。Tampermonkey 会按自身设置定期检查更新。

**每次发布新版本都必须递增脚本头部的 `@version`**（例如 `1.7.1` → `1.7.2`），并更新同一路径的 `xueqiu_stock_cube_finder.user.js`。

## 特性

- 按股票中文名称查询雪球相关组合并自动请求搜索结果分页
- 按仓位、累计收益、年化收益等排序
- 批量查询组合关停状态，永久缓存已关停组合
- 单条或批量获取最近主动调仓及核验当前真实持仓
- 配置股票搜索缓存有效期；持仓、调仓等数据采用独立缓存
- 导出 XLSX，宽屏表格和窄屏卡片布局

## 注意

雪球的股票搜索索引与组合实时持仓可能存在差异。仅以当前持仓核验结果判断当下是否持有。接口属于雪球网页服务，可能调整或限流。不要上传雪球 Cookie、access_token 或其他认证凭证。

本项目非雪球官方产品，也不构成投资建议。
