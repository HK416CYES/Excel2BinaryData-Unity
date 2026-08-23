# Excel2BinaryData-Unity

一个面向 Unity 的通用 Excel 配置表导出示例：在编辑器中读取 `.xlsx` / `.xls` 工作簿，生成 C# 数据结构类与容器类，并把配置数据写入二进制文件；运行时再将二进制表加载为可按主键查询的数据字典。

项目的核心特点是：**每个 Excel 工作簿通过固定的 `Config` 工作表声明数据表的读取规则**，不再要求所有工作表使用同一套固定表头行。

## 功能概览

- 支持一个 Excel 文件包含多张数据表。
- 支持 `int`、`float`、`bool`、`string` 四种字段类型。
- 自动生成每张表对应的 C# 数据类与 `Container` 容器类。
- 自动写入 `Assets/StreamingAssets/Binary/*.wang` 二进制表。
- 运行时通过 `BinaryDataManager` 将二进制数据还原为 `Dictionary<TKey, TData>`。
- 导出前全量校验；任意表有问题时停止整个导出，避免半成品文件。
- 错误信息包含 Excel 文件、工作表、行、列与原因。

## 目录说明

```text
Assets/
├─ ArtRes/Excel/                         # Excel 源文件
├─ Editor/Excel/
│  ├─ ExcelTool.cs                       # 编辑器导出工具
│  └─ ExcelDll/                          # Excel 读取依赖
├─ Scripts/ExcelData/
│  ├─ BinaryDataManager.cs               # 运行时二进制加载器
│  ├─ DataClass/                         # 自动生成的数据结构类
│  └─ Container/                         # 自动生成的字典容器类
└─ StreamingAssets/Binary/               # 自动生成的 .wang 二进制表
```

> `DataClass`、`Container` 与 `StreamingAssets/Binary` 下的内容均为导出产物。应修改 Excel 后重新导出，而不是手动维护这些生成文件。

## 快速开始

### 1. 放置 Excel 文件

将 `.xlsx` 或 `.xls` 文件放入：

```text
Assets/ArtRes/Excel/
```

每个工作簿都必须有一张名称严格为 `Config` 的工作表。它仅描述读取规则，不会生成数据类或二进制表。

### 2. 配置 `Config` 工作表

`Config` 的第 1 行可填写标题供人工阅读；程序从第 2 行开始，按固定 A 到 F 列读取：

| 列 | 建议标题 | 含义 |
|---|---|---|
| A | 工作表名 | 需要导出的数据工作表名称 |
| B | 数据起始行 | 第一条数据所在的 Excel 行号 |
| C | 数据起始列 | 第一个字段所在的 Excel 列号 |
| D | 变量名行 | 字段名所在的 Excel 行号 |
| E | 变量类型行 | 字段类型所在的 Excel 行号 |
| F | 主键列 | 主键字段所在的 Excel 列号 |

所有坐标均为 **Excel 的绝对坐标，且从 1 开始计数**。

例如，数据表的第 1 行是字段名、第 2 行是字段类型、第 5 行开始写数据，且第一列为主键：

| 工作表名 | 数据起始行 | 数据起始列 | 变量名行 | 变量类型行 | 主键列 |
|---|---:|---:|---:|---:|---:|
| TestInfo | 5 | 1 | 1 | 2 | 1 |

### 3. 编写数据工作表

字段名行从“数据起始列”开始连续读取，到第一个空字段名为止。字段名会被直接生成 C# 成员，因此必须是合法、非关键字且不重复的 C# 标识符。

字段类型行必须与字段名一一对应，仅支持以下精确拼写：

| 类型 | 说明 |
|---|---|
| `int` | 32 位整数 |
| `float` | 单精度浮点数 |
| `bool` | 布尔值，填写 `true` 或 `false` |
| `string` | UTF-8 字符串，可为空 |

数据从“数据起始行”开始读取，直到主键列首次出现空单元格。主键必须位于字段读取范围内，且所有主键值必须唯一。

### 4. 在 Unity 中导出

等待脚本编译完成后，在 Unity 顶部菜单选择：

```text
GameTool > GenerateExcel
```

成功时会提示导出数量，并生成：

- `Assets/Scripts/ExcelData/DataClass/<表名>.cs`
- `Assets/Scripts/ExcelData/Container/<表名>Container.cs`
- `Assets/StreamingAssets/Binary/<表名>.wang`

失败时不会写入新的生成文件；Console 和弹窗会报告具体位置。例如：

```text
文件[TestInfo.xlsx]，工作表[TestInfo]，行[2]，列[2]：字段类型[string ]不受支持，只允许int、float、bool、string。
```

## 运行时加载与访问

在游戏初始化阶段调用 `InitData()`：

```csharp
using UnityEngine;

public class GameDataBootstrap : MonoBehaviour
{
    private void Start()
    {
        BinaryDataManager.Instance.InitData();

        TestInfoContainer testTable =
            BinaryDataManager.Instance.GetTable<TestInfoContainer>();

        TestInfo test = testTable.dataDic[1];
        Debug.Log(tower.name);
    }
}
```

`InitData()` 当前显式加载 `TestInfo` 示例表。新增表后，需要在 `BinaryDataManager.InitData()` 中增加对应的加载调用：

```csharp
LoadTable<NewTableContainer, NewTable>();
```

然后通过以下方式获取容器：

```csharp
NewTableContainer table = BinaryDataManager.Instance.GetTable<NewTableContainer>();
```

## Excel 规则与边界

- 每个 Excel 必须有且仅应使用一张 `Config` 工作表。
- `Config` 中 A 到 F 列的每条记录必须完整；完全空白行会被跳过。
- 除 `Config` 外，所有**非空**工作表必须在 `Config` 中登记；完全空白工作表会被忽略。
- 同一工作表不能重复登记；`Config` 自身不能登记为数据表。
- 数据起始行必须位于字段名行、字段类型行之后。
- 主键第一次为空即视为数据结束；之后若仍存在字段数据，会被判定为主键断档。
- 表名会生成 C# 类名，字段名会生成 `public` 字段，因此二者都必须遵循 C# 命名规则。
- `string `、`int ` 等带空格的类型不是合法类型；请检查单元格末尾空格。

## 二进制文件位置与发布注意事项

配置二进制文件输出到：

```text
Assets/StreamingAssets/Binary/
```

`StreamingAssets` 会随构建产物一起发布，适合保存需要在运行时读取的原始二进制数据。当前加载器使用 `File.ReadAllBytes` 读取文件，因此该实现适合桌面平台及可直接访问 `StreamingAssets` 文件路径的平台；若目标平台需要通过 `UnityWebRequest` 访问 `StreamingAssets`，应为该平台补充对应的读取方式。

## 关于 `Save` / `Load`

`BinaryDataManager` 还提供了 `Save` 与 `Load<T>`，用于把普通对象保存到 `Application.persistentDataPath/Data/`。这与 Excel 导出的 `.wang` 表格式是两条独立流程。

该功能使用 `BinaryFormatter`。请勿反序列化来自网络、玩家上传或其他不可信来源的文件；如需面向外部输入，建议改用 JSON、MessagePack 或经过严格校验的自定义格式。

