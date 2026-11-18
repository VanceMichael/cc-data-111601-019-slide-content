# slide-content

课件页面可见文字与讲者注释的分层采集相关资料接收服务，面向授权来源的原件领取、解释和持久查询。

## 当前能力

已有 PPTX ZIP 部件读取，并按包内真实关系采集内容：页面自有文字、当前页实际显示的母版固定形状、页面对应的讲者注释三者分开，各自标明所属页、受众和来源原因。

`src/reader.mjs` 提供读取契约；`src/transport.mjs` 只访问配置来源的同源地址，拒绝隐式跳转；`src/store.mjs` 保存原件摘要和解释结果。`src/service.mjs` 是产品入口，`src/provider.mjs` 是独立 HTTP 联调来源，内容为无个人信息的本地样例。

采集遵循原件的实际关系，而不是按文件名猜测：

- 页序取自 `presentation.xml` 的 `sldIdLst` 与其 `slide` 关系；页序与部件文件名无关。
- slide → layout → master 沿各自的 `.rels` 解析；关系目标缺失时只记录问题并停止继承，不会借用包内另一张同名布局或母版。
- 页面自己的 `txBody` 即使为空也是显式文本，不会用版式或母版的占位提示填充。
- 母版中的非占位形状在 `showMasterSp`（默认开启）下计入该页学员可见文字；slide 或 layout 任一处关闭后该页不再计入。母版/版式的占位提示只作为定义保留。
- 讲者注释只来自该页 `notesSlide` 关系，始终标记为讲者内容，不混入学员正文。
- 隐藏页保留在页序中并单独标识，不因隐藏而丢失关系信息。

## 环境与运行

Node.js 22.22 至 22.x，无 npm 依赖，无外部运行服务或下载步骤。Linux、Windows 均使用以下命令；干净目录可直接运行。

```sh
npm ci --offline --ignore-scripts
npm test
npm run build
npm run demo
```

`npm run demo` 自动在本机回环端口启动来源与产品、领取一件资料、读取已保存结果，然后关闭服务并清理临时目录。它展示当前产品能力，不模拟尚不存在的能力。演示没有固定等待、外网访问或人工操作，通常数秒结束。扩展产品后可按实际能力更新演示路径；不要以打印预先保存的结果替代运行。

直接 `npm start` 即可启动产品，默认自动在临时回环端口创建本地合成来源，产品默认监听 127.0.0.1:4102。退出产品时同步关闭自带来源。需要联调独立来源时先运行 `npm run source`，再设置 SOURCE_URL=http://127.0.0.1:4101 运行 `npm start`；显式来源不会由产品关闭。DATA_DIR 可设置资料目录，PORT 可覆盖产品监听端口。构建产物位于 dist，启动、测试与演示均不需要外部服务或下载。

## 采集结果

`inspect()` 返回 `slides`（一页一项，按页序）和 `segments`（逐段文字）。每个 segment 都带 `slide`、`slidePart`、`speaker`、`audience`、`role`（slide/master/notes）、`part`、`reasonCode` 与可读的 `reason`，因此每段文字都能回答：属于哪一页、是否面向讲者、为何出现在该页。每页还提供 `slideText`、`masterText`、`notesText`、`learnerText`、`speakerText` 以及 `layoutPart`、`masterPart`、`notesPart`、`hidden` 和 `missing`。未显示的占位提示单独放在 `placeholders`；关系缺失、外部目标或越界目标记入 `issues`。

## 已有接口

- GET /health：服务状态。
- POST /intakes，JSON 为 {"path":"/primary"}：按现有 reader 接收单件原件；内容解释失败不登记资料。
- GET /records：持久接收目录；同 source URL、同摘要重领返回原 id。
- GET /objects/{digest}：读取已保存的原始字节。

这四项是兼容接口，后续扩展可以新增业务入口，但不能把 GET 请求变为写操作。收件记录当前只包含单件 raw 对象，不代表关联资料已完整。

## 来源与资料

交换范围、字段和边界见 [来源约定](docs/upstream.md)，样例资源清单可访问来源的 /catalog。`fixtures/upstream.json` 描述 HTTP 资源和阶段变体，不包含产品实现。GET /_requests 可查实际领取请求；POST /_phase 接收 {"phase":1} 切换已定义的来源状态，切回 0 恢复。所有资料都限于公开约定或本地合成数据，不向真实业务系统写入。

## 测试与构建

`npm test` 覆盖读取器分层采集（页序、母版显示开关、占位提示、讲者注释、缺失关系、复用布局）、畸形输入、真实 HTTP 领取、原件一致性、持久重开、同源边界和范围响应；领域读取行为在 test/domain.test.mjs，合成 PPTX 构造器在 test-support/pptx-fixture.mjs。`npm run build` 对运行模块逐个执行语法校验并复制到 dist。`docker build -t slide-content:local .` 可创建产品运行镜像；镜像默认运行实际演示，不含远程业务凭据。
