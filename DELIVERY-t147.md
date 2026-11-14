# 读取容错契约重做说明（t147-exceljs）

## 背景

异构系统导出的 xlsx 包内部件命名与关系引用经常不规范，联调中发现两类互相矛盾的极端行为：

1. **静默丢数据**：工作表部件条目名大小写与常规命名不一致（如 `xl/Worksheets/sheet1.xml`）时，条目分发的正则与 `switch` 均区分大小写，该条目落入未匹配分支被静默跳过，工作簿目录（`xl/workbook.xml`）又通过精确路径在 `worksheetHash` 中查找，结果工作表被整体丢弃——读取正常返回但**零个工作表、零告警**。
2. **裸崩**：工作表关系件（`sheetN.xml.rels`）悬空指向包内不存在的批注部件时，`WorkSheetXform.reconcile` 直接解引用 `options.comments[rel.Target].comments`，抛出未受控的 `TypeError: Cannot read properties of undefined (reading 'comments')`，整个读取崩溃。

## 容错契约（三类异常的受控语义）

重做后，读取路径上的异常归为三类，统一由新增的 `lib/xlsx/read-warnings.js`（`ReadWarnings` 收集器）记录：

| 异常类别 | 告警 code | 受控语义 | 设计理由 |
| --- | --- | --- | --- |
| **未识别条目**（zip 内存在读取器不认识的部件路径） | `unrecognised-entry` | 跳过该条目，其余部件正常解析，读取成功完成并告警 | 包整体仍可解析，丢弃一个未知部件不应让读取失败；但"静默吞掉"会掩盖生产方的结构问题，因此必须留痕 |
| **部件缺失**（工作簿目录/关系引用的部件在包内不存在） | `missing-part` | 跳过受影响的单项功能（如某张工作表），其余工作簿内容继续加载，读取成功完成并告警 | 与"未识别条目"区分：这里有明确的引用方期待该部件，属于真实的数据降级，调用方需要知道哪张表丢了 |
| **悬空关系件**（关系 Target 无法解析到已解析部件，或关系 Id 本身缺失） | `dangling-relationship` | 只降级该关系承载的单项特性（批注/VML/超链接/图片/表格），工作表本体与单元格数据完整保留，读取成功完成并告警 | 关系是部件间的增强特性，其断裂不应升级为整表/整包失败；按特性粒度降级可最大程度保住数据 |

### 关键取舍

- **降级而非抛错**：三类异常都让 `load/readFile` 正常 resolve，而不是 reject。理由是这三类都属于"包结构不规范但数据可部分恢复"，硬失败会把可挽救的数据一并锁死；真正的内容损坏（如 XML 解析错误）仍维持原有的 reject 语义不变（`invalid-xml.xlsx` 用例保持通过）。
- **正常文件零告警**：所有标准 OOXML 布局但本读取器不消费的部件（`[Content_Types].xml`、`docProps/*`、`calcChain.xml`、`charts/*`、`chartsheets/*`、`pivotTables/*`、`printerSettings/*`、`customXml/*` 等）保持**静默跳过**，不产生告警。只有落在已知布局之外的真正未知条目才告警，避免噪声淹没真实问题。
- **大小写折叠**：条目分发的固定名匹配与动态部件正则统一改为大小写不敏感；`worksheetHash` 查找在精确匹配失败后回退到大小写折叠匹配。折叠事件记 `entry-name-case` 告警（按条目名去重，避免分发与协调两处重复记录）。
- **chartsheet 保持静默**：`chartsheets/sheet1.xml` 这类是 OOXML 合法部件、本库明确不建模，按既有约定继续静默忽略，不计入 `missing-part`。
- **缺工作簿部件语义保持**：不含 `xl/workbook.xml` 的包（既有 `missing-bits.xlsx` 用例）继续 resolve 为零工作表、零告警，这是已确立的集成契约，不新增告警。
- **可选部件缺失保持静默**：`styles.xml`、`sharedStrings.xml` 等可选部件历来容忍（对应 model 字段为 `undefined`，协调链逐处防御），不制造告警，保证常规包零告警。

## 告警的可编程读取

每次读取后，告警收集器挂在 xlsx 子系统实例上：

```js
const workbook = await new ExcelJS.Workbook().xlsx.load(buffer);
const warnings = workbook.xlsx.readWarnings;

warnings.count;          // 告警数量
warnings.toArray();      // 告警对象快照（副本，外部改动不影响内部状态）
for (const w of warnings) { ... }  // 可迭代
```

告警对象统一形状 `{code, message, ...details}`，`details` 携带定位信息（`entryName`/`canonicalName`/`partName`/`source`/`relId`/`target`/`relType`/`reason`/`sheetName` 等），便于程序化分流处理。正常读取时 `count === 0`。

## 主要改动

- `lib/xlsx/read-warnings.js`（新增）：`ReadWarnings` 告警收集器与四类 code 的工厂方法。
- `lib/xlsx/xlsx.js`：
  - `load()` 的 `switch` 分发改成"固定名分发器 + 动态部件正则表 + 未识别条目兜底"，全部正则加 `i` 标志；媒体/主题的流类型选择同步改为大小写不敏感。
  - 标准但不消费的部件走静默白名单，其余未识别条目记 `unrecognised-entry`。
  - `reconcile()` 把 `readWarnings` 传入工作表协调；读取结束后暴露为 `this.readWarnings`。
- `lib/xlsx/xform/book/workbook-xform.js`：`reconcile` 中工作表查找支持大小写折叠；目录引用的关系缺失记 `dangling-relationship`、目标部件缺失记 `missing-part`；chartsheet 目标保持静默。
- `lib/xlsx/xform/sheet/worksheet-xform.js`：消除批注、VML、超链接、drawing、背景图、表格六类关系解引用的裸 `TypeError`，逐类受控降级并记 `dangling-relationship`（含 `reason` 区分子因）。
- `lib/xlsx/xform/sheet/cell-xform.js`：`model.value.richText` 解引用加空值防御（缺共享字符串等场景）。
- `spec/integration/workbook-xlsx-malformed-package.spec.js`（新增）：坏包形态矩阵对照测试，12 个用例覆盖正常包零告警、目录/文件名大小写折叠、批注/VML 悬空、未识别条目、标准部件静默、缺工作表部件、缺关系、缺工作簿部件、告警收集器快照语义。

## 不回退验证

- 单元测试：**883 passing / 1 pending**，与改动前完全一致。
- 集成测试：新增 12 个矩阵用例全部通过；既有通过数不降。全量运行残留 3 个失败（issue 1328、pr 1431、`WorkbookReader > Serialise > xlsx file`），经 `git stash` 对照确认在**原始快照上同样存在**，且第三个失败堆栈位于流式读取器 `lib/stream/xlsx/workbook-reader.js`，不在本次改动的非流式 `lib/xlsx/` 路径内。
- 真实文件（gold/images/formulas/chart-sheet/test-issue-880/fibonacci/1904 等）读取结果与告警均符合预期：正常包零告警、数据完整。
- `invalid-xml.xlsx`（XML 内容损坏）仍按原契约 reject，未被降级逻辑吞掉。
- ESLint 对全部改动文件零 error、零 warning。
