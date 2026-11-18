# 来源约定

## 产品交换范围

OOXML PresentationML 子集包含 presentation 的页顺序关系、slide/layout/master 的关系与 p:spTree/p:sp、p:ph(type,idx)、a:p/a:r/a:t、notesSlide。沿 package relationship Type 判角色，Target 按当前部件相对 URI 解析。页面自有 txBody 即便空也表示显式文本，不能用布局或母版占位提示填充；占位符可继承样式不代表继承提示文字。普通非占位母版形状在 showMasterSp=true（默认）时作为独立来源的固定文字收集，layout/slide 关闭对应显示后不加入该页；本项目不涉及群组层级、动画、页脚字段计算或形状遮挡。版式/母版中的 placeholder 提示保留为定义而非 learner 正文。hidden slide 保留且独立标识。notes 保持 speaker 角色。无须渲染或 OCR，不要求图表数据和 SmartArt。

## 采集结果分层

`inspect()` 返回 `slides`（按 presentation 页序，一项一页）与 `segments`（逐段文字）。每段都带 `slide`、`slidePart`、`speaker`、`audience`、`role`、`part`、`reasonCode` 和可读的 `reason`，用于说明这段文字属于哪一页、是否面向讲者、以及为什么出现在该页：

- `role: "slide"`：页面部件自身的显式文字，`speaker: false`。
- `role: "master"`：当前页所用母版中的非占位形状文字，仅在 slide 与 layout 都未关闭 `showMasterSp` 时收集。
- `role: "notes"`：该页通过 `notesSlide` 关系关联的讲者注释，`speaker: true`，不进入 `learnerText`。

每页同时提供 `slideText`、`masterText`、`notesText`、`learnerText`（slide + master）和 `speakerText`（notes）以及 `layoutPart`/`masterPart`/`notesPart`。`hidden` 独立标识隐藏页，隐藏页仍保留页序与关系，只是其显示来源不计正文。`placeholders` 单独列出 layout/master 中未显示的占位提示定义（含 `referencedBy`），不会混入正文。`issues` 记录缺失、外部或越界的关系目标。

关系解析只认当前部件 `.rels` 中声明的 Target。目标部件不存在时记入 `issues` 并停止继承，不会用包内另一张同名或相似命名的 layout/master 补出内容。没有起点的 `a:fld`（页码、日期等计算字段）与图片、图表、SmartArt 文字不在本子集内。

## 当前来源路由

- `/primary`：application/vnd.openxmlformats-officedocument.presentationml.presentation
- `/missing-rel`：application/octet-stream

`/catalog` 提供入口及业务选择。资源字节支持 GET/HEAD、单一 bytes 范围、ETag 和 If-Match，响应不会跨来源重定向。`/_phase` 为联调控制路径，不是产品数据来源，不应由产品自行切换阶段。

## 依据和范围说明

- https://ecma-international.org/publications-and-standards/standards/ecma-376/

以上是格式或业务词汇的公开依据；本项目明确限定的交换子集以本文及实际资源为准，未要求完整标准实现。元数据和未支持声明应保留可定位的来源，不把未知转换成已确认的空结果。HTTP 来源状态、原件结构和产品接收完整性是不同层次；现有通用收件接口成功只表示该单件已保存。
