# 来源约定

## 产品交换范围

OOXML PresentationML 子集包含 presentation 的页顺序关系、slide/layout/master 的关系与 p:spTree/p:sp、p:ph(type,idx)、a:p/a:r/a:t、notesSlide。沿 package relationship Type 判角色，Target 按当前部件相对 URI 解析。页面自有 txBody 即便空也表示显式文本，不能用布局或母版占位提示填充；占位符可继承样式不代表继承提示文字。普通非占位母版形状在 showMasterSp=true（默认）时作为独立来源的固定文字收集，layout/slide 关闭对应显示后不加入该页；本项目不涉及群组层级、动画、页脚字段计算或形状遮挡。版式/母版中的 placeholder 提示保留为定义而非 learner 正文。hidden slide 保留且独立标识。notes 保持 speaker 角色。无须渲染或 OCR，不要求图表数据和 SmartArt。

## 当前来源路由

- `/primary`：application/vnd.openxmlformats-officedocument.presentationml.presentation
- `/missing-rel`：application/octet-stream

`/catalog` 提供入口及业务选择。资源字节支持 GET/HEAD、单一 bytes 范围、ETag 和 If-Match，响应不会跨来源重定向。`/_phase` 为联调控制路径，不是产品数据来源，不应由产品自行切换阶段。

## 依据和范围说明

- https://ecma-international.org/publications-and-standards/standards/ecma-376/

以上是格式或业务词汇的公开依据；本项目明确限定的交换子集以本文及实际资源为准，未要求完整标准实现。元数据和未支持声明应保留可定位的来源，不把未知转换成已确认的空结果。HTTP 来源状态、原件结构和产品接收完整性是不同层次；现有通用收件接口成功只表示该单件已保存。
