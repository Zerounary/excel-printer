# excel-printer

基于 `config.json` 的**Excel 打印模板生成器**。

- 输入：一个 JSON 配置文件（默认 `config.json`），或一个包含多个 JSON 的文件夹
- 输出：单个 `.xlsx`（默认 `output.xlsx`）或批量输出到某个目录（默认 `out/`）

## 安装与生成

```bash
npm install
npm run generate
```

等价于：

```bash
# 单个 config
node src/index.js --config config.json --out output.xlsx
# 批量模式（默认输出到 out/）
node src/index.js --config-dir configs --out-dir out
```

也可以通过可执行命令运行（安装为依赖后可用）：

```bash
excel-printer --config config.json --out output.xlsx
excel-printer --config-dir configs --out-dir out
```

> - `--config-dir` 提供时，会遍历该目录下所有 `.json` 文件批量生成 Excel。此时 `--config/--out` 会被忽略。
> - 目录模式下，系统会把 **JSON 文件名当作模板**：例如把 `{{today}}-{{company.name}}-发货单.json` 解析成 `2025-01-01-某某公司-发货单.xlsx`。文件名模板可使用 config `variables` 里的任意字段，另内置 `{{file.name}}`、`{{file.baseName}}`、`{{today}}` 以及 `{{date.year}}`/`month`/`day`。非法文件字符会自动替换，并补全 `.xlsx` 后缀。

## 作为库使用（推荐）

本项目同时提供可复用的生成 API（ESM）。

默认导出为 **可在浏览器/Vite 环境复用的 core 入口**（不包含文件读写）：

```js
import { generateWorkbookFromConfig } from 'excel-printer';

const workbook = generateWorkbookFromConfig(config);
// Node 环境：workbook.xlsx.writeFile('output.xlsx')
// 浏览器环境：workbook.xlsx.writeBuffer() -> Blob -> 下载
```

Node 环境如果需要“读取 config 文件 + 直接写出 xlsx 文件”，请使用 Node 入口：

```js
import { generateXlsxFileFromConfigFile } from 'excel-printer/node';

await generateXlsxFileFromConfigFile({
  configPath: 'config.json',
  outPath: 'output.xlsx',
});
```

可用 API：

- core（浏览器/Node 通用）：`generateWorkbookFromConfig(config)`
- node（仅 Node）：
  - `generateXlsxFileFromConfig(config, outPath)`
  - `generateXlsxFileFromConfigFile({ configPath, outPath })`
  - `generateXlsxFilesFromConfigDir({ configDir, outDir })` —— 批量读取目录内所有 JSON，并把结果写到 `outDir`（默认 `out/`），输出文件名自动根据 JSON 文件名模板计算

## JSON 文件名模板规则

在批量模式（`--config-dir` 或 `generateXlsxFilesFromConfigDir`）下，JSON 文件名（去掉 `.json`）会被当作模板字符串，支持与 config 内相同的 `{{path.to.var}}` 语法：

- `{{file.name}}` / `{{file.baseName}}`：原始文件名与去掉扩展名后的部分。
- `{{today}}`：当前日期，格式 `YYYY-MM-DD`。
- `{{date.year}}` / `{{date.month}}` / `{{date.day}}`：当前年月日。
- 以及配置里声明的任意 `variables`。

示例：`files/{{today}}-{{company.name}}-发货单.json` 会输出到 `out/2026-01-05-示例供应链有限公司-发货单.xlsx`。

## 变量替换（模板）

为了方便你“先拿到数据，再快速组装成 config 并生成 Excel”，支持在配置中声明变量，并在任意字段里做替换。

### 1) variables

顶层可声明：

```json
{
  "variables": {
    "company": "XX农业发展有限公司",
    "date": "2026-01-05",
    "tableRows": []
  }
}
```

也兼容 `vars`。

### 2) 字符串占位符替换：{{path.to.var}}

示例：

```json
{ "type": "title", "value": "{{company}}送货单" }
```

### 3) 非字符串替换：{"$var": "path.to.var"}

当你要替换整个数组/对象/数值时，用 `$var`：

```json
{
  "type": "table",
  "headers": ["商品名称", "数量"],
  "rows": { "$var": "tableRows" }
}
```

## 配置字段命名规范（value）

- `title` 与 `text` 块统一使用 `value`
- 旧字段 `val` 仍兼容（会自动归一化为 `value`）

## config.json 总体结构

推荐使用“模板 + 实例”结构：

```json
{
  "style": {},
  "variables": {
    "shared": {},
    "sheets": []
  },
  "sheetsTemplates": []
}
```

- `style`（可选）
  - **全局默认样式**。会被各 block 的 `style` 覆盖。
- `variables`
  - 全局变量（`shared` 示例仅表示你可自定义结构）
  - `sheets`：**实例数组**，每个元素表示一个真实 sheet，需要指定使用哪个模板
- `sheetsTemplates`
  - 模板数组，每个元素描述一个可复用的 sheet 布局（原来的 `sheets` 就是现在的模板内容）

### sheetsTemplates[] 结构（模板）

```json
{
  "id": "delivery",
  "name": "送货单模板",
  "paper": "A4",
  "maxColumns": 6,
  "rows": []
}
```

- `id`（推荐）：模板唯一标识，`variables.sheets[].template` 会引用它
- 其他字段（`name/maxColumns/rows/...`）与旧 `sheets[]` 结构一致

### variables.sheets[] 结构（实例）

```json
{
  "template": "delivery",
  "name": "2026-01-05 送货单",
  "variables": {
    "delivery": {
      "date1": "2026-01-05",
      "date2": "2026-01-06"
    }
  }
}
```

- `template` / `sheetsTemplate` / `templateId`
  - 选择要复用的模板（匹配 `id` 或 `name`）
- `name`
  - 真实 sheet 名称（若省略，将回落到模板的 `name`）
- `variables` / `vars` / `data`
  - 该 sheet 的私有变量，会与全局 `variables` 深度合并
- 你也可以直接在实例对象上写与模板同名的字段（例如 `paper`、`rows`），用于覆盖模板

> 额外内置变量：渲染时会注入 `sheet = { name, index, template }`，因此模板里可以写 `{{sheet.name}}` 等。

### 兼容模式：直接使用 sheets[]

仍然支持旧写法：

```json
{
  "sheets": [
    {
      "name": "打印",
      "maxColumns": 6,
      "rows": []
    }
  ]
}
```

以及更老的单-sheet 写法：

```json
{
  "maxColumns": 6,
  "rows": []
}
```

这两种都会被自动转换为单个 sheet。

## 项目架构 / 目录说明

```
src/
├── index.js        # CLI：解析参数 -> 调用库 generate.js
├── generate.js     # 核心库入口：多 sheet 渲染、写文件、兼容旧配置
├── cli.js          # 命令行参数解析（--config / --out）
├── utils.js        # 通用辅助：maxColumns 归一化、列名转换
├── layout.js       # 布局相关：列宽/分页设置、合并单元格、行高估算
├── styles.js       # 样式工具：样式合并/应用、默认边框
└── renderers.js    # 渲染器：title / text / form / table 的绘制逻辑
```

- **CLI (`index.js`)**：仅负责命令行调用，不承载核心逻辑。
- **库入口 (`generate.js`)**：核心生成逻辑（建议其他项目直接引用这里导出的 API）。
- **渲染器 (`renderers.js`)**：针对每种 block 类型封装绘制逻辑，便于单独调整。
- **layout / styles**：与内容无关的横切逻辑单独存放，方便复用与测试。
- **cli / utils**：聚合纯函数工具，避免 `index.js` 出现太多杂项逻辑。

> 若要扩展新的 block 类型或渲染策略，只需在 `renderers.js` 新增对应函数，并在 `index.js` 的遍历中接入即可。

## 默认行为

- 会根据 `rows` 的顺序自动向下排版（无需你手动写单元格坐标）。
- 渲染完成后，会自动设置打印区域（print area）为 `A1` 到内容最后一行。
- **每个数据格默认都有细边框**：对实际打印区域内所有单元格进行边框补齐。
- 默认不会在 `title/form/table/text` 之间自动插入空白行；如需间距请使用 `space-row`。

## 样式（style）配置

你可以在 `config.json` 中为不同层级配置 `style`，用于控制：

- 字体：字体、字号、加粗、斜体等
- 对齐：水平/垂直、是否换行
- 边框：细线/粗线/自定义
- 填充：背景色
- 数字格式：`numFmt`

### style 对象格式

`style` 的字段直接映射到 `exceljs` 的 cell 属性：

```json
{
  "font": { "name": "宋体", "size": 11, "bold": false, "italic": false },
  "alignment": { "horizontal": "left", "vertical": "middle", "wrapText": true },
  "border": {
    "top": { "style": "thin" },
    "left": { "style": "thin" },
    "bottom": { "style": "thin" },
    "right": { "style": "thin" }
  },
  "fill": {
    "type": "pattern",
    "pattern": "solid",
    "fgColor": { "argb": "FFEFEFEF" }
  },
  "numFmt": "0.00"
}
```

> 说明：如果你不填 `style`，会使用程序内置的默认样式（保持现状）。

## rows 支持的 block 类型

### 1) 标题 title

```json
{
  "type": "title",
  "val": "送货单",
  "style": {
    "font": { "size": 18, "bold": true },
    "alignment": { "horizontal": "center" }
  }
}
```

- 标题默认会合并为整行（从第 1 列合并到 `maxColumns`）。

### 2) 文本 text

```json
{
  "type": "text",
  "value": "注：请核对数量和质量...",
  "style": {
    "font": { "italic": true },
    "alignment": { "horizontal": "left", "wrapText": true }
  }
}
```

- `text` 默认也是整行合并。

### 3) 表单 form

```json
{
  "type": "form",
  "style": { "font": { "size": 11 } },
  "fieldStyle": { "alignment": { "horizontal": "left" } },
  "fields": [
    {
      "type": "text",
      "label": "收货单位",
      "value": "xxx",
      "style": { "font": { "bold": true } }
    },
    {
      "type": "text",
      "label": "日期",
      "value": "2026-01-05"
    }
  ]
}
```

- 当前实现：一行默认放 2 个 field（自动计算分配列宽区间，并做合并）。
- 样式合并优先级：`config.style` -> `block.style` -> `block.fieldStyle` -> `field.style`

### 4) 表格 table

```json
{
  "type": "table",
  "headers": ["商品名称", "分类", "单位", "订购数量", "实收数量", "备注"],
  "rows": [
    ["一级乡厨房大米25公斤/袋", "普通大米", "袋", "663", "", ""],
    ["...", "...", "...", "...", "...", "..."]
  ],

  "style": { "font": { "size": 11 } },
  "headerStyle": {
    "font": { "bold": true },
    "alignment": { "horizontal": "center" },
    "fill": {
      "type": "pattern",
      "pattern": "solid",
      "fgColor": { "argb": "FFF3F3F3" }
    }
  },
  "bodyStyle": {
    "alignment": { "wrapText": true }
  },

  "columnStyles": [
    { "alignment": { "horizontal": "left" } },
    { "alignment": { "horizontal": "center" } },
    { "alignment": { "horizontal": "center" } },
    { "alignment": { "horizontal": "right" } },
    { "alignment": { "horizontal": "right" } },
    { "alignment": { "horizontal": "left" } }
  ],

  "rowStyles": [
    { "font": { "bold": false } },
    { "font": { "italic": false } }
  ],

  "cellStyles": {
    "0,1": { "font": { "bold": true } },
    "1,4": { "font": { "bold": true }, "alignment": { "horizontal": "right" } }
  }
}
```

#### table 的样式覆盖层级

- **表头单元格**（header 行）：
  - `config.style` -> `block.style` -> `block.headerStyle` -> `block.columnStyles[colIndex]` -> `block.cellStyles["0,col"]`
- **数据单元格**（body 行）：
  - `config.style` -> `block.style` -> `block.bodyStyle` -> `block.columnStyles[colIndex]` -> `block.rowStyles[rowIndex]` -> `block.cellStyles["row,col"]`

#### cellStyles 的 key 规则

- key 格式：`"rowIndex,colIndex"`
- `rowIndex`：
  - `0` 表示表头行
  - `1..N` 表示第 1..N 条数据行
- `colIndex`：
  - 从 `1` 开始，到 `maxColumns`

### 5) 空白行 space-row

用于手动控制 block 之间的留白。

特性：

- 会把该行从第 1 列合并到 `maxColumns`
- 该行会被标记为“无边框行”，不会被默认边框逻辑补齐，因此不会出现竖线/条纹

配置示例：

```json
{
  "type": "space-row",
  "count": 1,
  "height": 18
}
```

- `count`（可选）
  - 空白行数量，默认 `1`
- `height`（可选）
  - 行高（Excel 行高单位），不填则使用默认行高

## 常见问题

### 如何让某些格子更“像打印表单”？

- 可以通过 `style.border` 为特定格子设置更粗的底边/外边框。
- 可以对签字区 field 设定更大的字号、或使用 `border.bottom` 模拟下划线。

### 能否去掉边框？

- 当前默认行为是“打印区域内每个格子都有边框”。如果你确实需要“无边框区域”，需要在代码里增加一个开关（例如 `config.defaultBorder: false`）。
