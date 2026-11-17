# slide-content

课件中页面可见文字与讲者注释的分层采集相关资料接收服务，面向授权来源的原件领取、解释和持久查询。

## 当前能力

PPTX ZIP 部件读取、presentation 页顺序关系解析，以及页面可见文字与讲者注释的分层采集。

`src/reader.mjs` 沿 package relationship 解析 slide→layout→master 与 notesSlide 的实际关系链（Target 按声明部件相对 URI 解析，角色按 relationship Type 判定），按 presentation 的 sldIdLst 确定页顺序。每页采集结果分层记录：

- `texts`：收进检索库的文字，逐段标注 `audience`（learner/speaker）、`origin`（slide/layout/master/notes）与来源部件。页面自有 txBody 即便空也保持显式空，不用版式或母版占位提示填充；母版非占位固定形状仅在页面与版式均未关闭 showMasterSp 时加入该页。
- `placeholderDefinitions`：版式/母版占位符的提示文字，保留为定义，不作为 learner 正文。
- `relations`/`missing`：该页声明的关联部件及其中包内缺失者；缺失时不借另一张同名布局补全，该页仅记录实际存在的来源。
- `visible`：show="0" 的隐藏页保留且标识为 false。

`src/transport.mjs` 只访问配置来源的同源地址，拒绝隐式跳转；`src/store.mjs` 保存原件摘要和解释结果。`src/service.mjs` 是产品入口，`src/provider.mjs` 是独立 HTTP 联调来源，内容为无个人信息的本地样例。

## 环境与运行

Node.js 22.22 至 22.x，无 npm 依赖，无外部运行服务或下载步骤。Linux、Windows 均使用以下命令；干净目录可直接运行。

```sh
npm ci --offline --ignore-scripts
npm test
npm run build
npm run demo
```

`npm run demo` 自动在本机回环端口启动来源与产品、领取一件资料、读取已保存结果，然后关闭服务并清理临时目录。它展示当前产品能力，不模拟尚不存在的能力。演示没有固定等待、外网访问或人工操作，通常数秒结束。扩展产品后可按实际能力更新演示路径；不要以打印预先保存的结果替代运行。

分开运行时，先 `npm run source`，来源监听 127.0.0.1:4101；在另一终端设置 SOURCE_URL=http://127.0.0.1:4101 后运行 `npm start`，产品默认 4102。DATA_DIR 可设置资料目录，PORT 可分别覆盖监听端口。构建产物位于 dist，当前脚本仅使用 Node 内置模块。

## 已有接口

- GET /health：服务状态。
- POST /intakes，JSON 为 {"path":"/primary"}：按现有 reader 接收单件原件；内容解释失败不登记资料。
- GET /records：持久接收目录；同 source URL、同摘要重领返回原 id。
- GET /objects/{digest}：读取已保存的原始字节。

这四项是兼容接口，后续扩展可以新增业务入口，但不能把 GET 请求变为写操作。收件记录的 `parsed` 字段保存上述分层采集结果；记录仍对应单件 raw 对象，不代表关联资料已完整。

## 来源与资料

交换范围、字段和边界见 [来源约定](docs/upstream.md)，样例资源清单可访问来源的 /catalog。`fixtures/upstream.json` 描述 HTTP 资源和阶段变体，不包含产品实现。GET /_requests 可查实际领取请求；POST /_phase 接收 {"phase":1} 切换已定义的来源状态，切回 0 恢复。所有资料都限于公开约定或本地合成数据，不向真实业务系统写入。

## 测试与构建

`npm test` 覆盖现有读取器、畸形输入、真实 HTTP 领取、原件一致性、持久重开、同源边界和范围响应；领域读取行为在 test/domain.test.mjs。`npm run build` 对运行模块逐个执行语法校验并复制到 dist。`docker build -t slide-content:local .` 可创建产品运行镜像；镜像默认运行实际演示，不含远程业务凭据。
