---
name: excel-dna
description: 生成基于 Excel-DNA 的 Excel 插件（.xll 文件）。当用户需要创建 Excel 插件、Excel 加载项、Excel Ribbon 功能区、Excel 自定义函数、Excel 宏插件，或提及 Excel-DNA、.xll、.dna 等关键词时使用此技能。支持生成 .dna 配置文件、.cs 代码文件，并自动打包发布可用的 .xll 插件。
allowed-tools: Read, Write, Bash
---

# Excel-DNA 插件生成技能

## 概述

本技能根据用户自然语言描述，自动生成完整的 Excel-DNA 插件项目，包括 .dna 配置文件、.cs 代码文件，并自动复制重命名 ExcelDna64.xll 文件，生成可直接在 Excel 中加载的 .xll 插件。

## 工作流程

按以下步骤执行：

1. **解析用户需求**：从用户描述中提取项目名称（projectName）、功能区名称、按钮列表及对应功能
2. **设计 CustomUI**：根据按钮和功能设计 Ribbon XML
3. **生成 .dna 文件**：创建同名 .dna 配置文件
4. **生成 .cs 文件**：根据复杂度决定单文件或多文件拆分
5. **复制 XLL 文件**：将 ExcelDna64.xll 复制到输出目录并重命名为 {projectName}.xll
6. **输出使用说明**：告知用户如何加载和使用插件

## 输出目录

默认在项目根目录下创建 `{projectName}/` 子目录，将所有文件输出到该目录。
XLL 模板文件路径：`${CODEBUDDY_SKILL_DIR}/../../../ExcelDna-1.9.0/ExcelDna64.xll`

## .dna 文件格式规范

生成名为 `{projectName}.dna` 的 XML 文件，格式如下：

```xml
<DnaLibrary Name="{ProjectName}" RuntimeVersion="v4.0">
  <Project Language="C#">
    <Reference Name="System.Drawing" />
    <Reference Name="System.Windows.Forms" />
    <Reference Name="Microsoft.Office.Interop.Excel" />
    <SourceItem Path="Main.cs" />
    <!-- 如有多个 .cs 文件，每个文件一个 SourceItem -->
  </Project>
  <CustomUI>
    <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='RibbonLoad'>
      <ribbon>
        <tabs>
          <tab id='CustomTab' label='{TabLabel}'>
            <group id='SampleGroup' label='{GroupLabel}'>
              <button id='Button1' label='{ButtonLabel}' size='large' onAction='{ButtonId}Click' />
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>
  </CustomUI>
</DnaLibrary>
```

关键规则：
- Root 元素必须是 `DnaLibrary`，设置 `Name` 和 `RuntimeVersion="v4.0"`
- `Project` 元素 `Language="C#"`
- `System`、`System.Collections.Generic`、`System.Linq`、`System.Text` 已内置，无需引用
- 可选引用：`System.Drawing`、`System.Windows.Forms`、`Microsoft.Office.Interop.Excel`
- 每个 .cs 文件对应一个 `SourceItem` 元素，`Path` 为文件名
- `CustomUI` 中的 `onLoad` 回调和按钮的 `onAction` 回调必须在 C# 代码中实现

## .cs 文件格式规范

### 基础模板

```csharp
using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
// Application 与 System.Windows.Forms.Application 冲突，使用别名
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

public class MyRibbon : ExcelRibbon
{
    // CustomUI onLoad 回调
    public void RibbonLoad(IRibbonUI ribbon)
    {
        // 初始化逻辑
    }

    // 按钮回调 - 必须为 public
    public void Button1Click(IRibbonControl control)
    {
        // 获取 Excel Application 对象
        ExcelApp app = (ExcelApp)ExcelDnaUtil.Application;
        try
        {
            // 操作 Excel
            Workbook wb = (Workbook)app.ActiveWorkbook;
            Worksheet ws = (Worksheet)wb.ActiveSheet;
            Range range = (Range)ws.Selection;
            
            // 业务逻辑
            string text = Convert.ToString(range.Value);
            MessageBox.Show(text);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error: " + ex.Message);
        }
    }
}
```

关键规则：
- 必须引用 `ExcelDna.Integration` 和 `ExcelDna.Integration.CustomUI`
- 使用 `using Microsoft.Office.Interop.Excel;` 引入 Excel 类型（`Workbook`、`Worksheet`、`Range` 等）
- 使用 `using ExcelApp = Microsoft.Office.Interop.Excel.Application;` 别名避免与 `System.Windows.Forms.Application` 冲突
- 必须继承 `ExcelRibbon` 类
- CustomUI 中定义的事件回调必须为 `public`
- 必须通过 `ExcelDnaUtil.Application` 获取 `Application` 对象，并用 `(ExcelApp)` 转换
- 所有 COM 对象必须显式类型转换：`(Workbook)app.ActiveWorkbook`、`(Worksheet)wb.ActiveSheet`
- 从 Range 取值使用 `Convert.ToString()`、`Convert.ToDouble()` 等方法进行类型转换

## C# 4.0 语法约束

生成的代码必须严格使用 C# 4.0 语法，禁止使用以下特性：

| 禁止使用的语法 | 说明 | 正确替代方案 |
|---|---|---|
| `$"text {var}"` | $字符串插值 | `string.Format("text {0}", var)` |
| `?.` 空条件运算符 | Null-conditional operator | `obj != null ? obj.Value : null` |
| `??=` Null 合并赋值 | Null-coalescing assignment | `if (obj == null) obj = defaultValue;` |
| `async/await` | 异步编程 | 使用回调或同步方法 |
| `nameof()` | 名称of表达式 | 直接使用字符串字面量 |
| 自动属性初始化 `public int X { get; set; } = 1;` | C# 6.0+ | 在构造函数中初始化 |
| 字典初始化器 `new Dictionary<K,V> { [key] = value }` | C# 6.0+ | 使用 `Add()` 方法 |
| Lambda 多行（某些场景） | 受限 | 使用完整方法定义 |
| `out var` 变量声明 | C# 7.0+ | 先声明变量再传入 |
| 元组语法 `(int, string)` | C# 7.0+ | 使用 `Tuple<T1,T2>` 或自定义类 |
| 局部函数 | C# 7.0+ | 使用私有方法 |
| `switch` 表达式 | C# 8.0+ | 使用传统 `switch` 语句 |
| 顶级语句 | C# 9.0+ | 必须有 `class` 和 `Main` 方法（但 ExcelDNA 不需要 Main） |

## 文件拆分规则

根据功能复杂度决定 .cs 文件数量：

| 复杂度 | 条件 | 文件结构 |
|---|---|---|
| 简单 | 1-2个按钮，无独立业务逻辑 | 单一 `Main.cs` |
| 中等 | 3-5个按钮，或有可复用的辅助方法 | `Ribbon.cs` + `Helpers.cs` |
| 复杂 | 5+个按钮，或有独立的业务模块 | `Ribbon.cs` + `Functions.cs` + `Helpers.cs` |

拆分原则：
- `Ribbon.cs`：继承 `ExcelRibbon`，实现 CustomUI 事件回调
- `Functions.cs`：实现自定义 Excel 函数（使用 `[ExcelFunction]` 特性）
- `Helpers.cs`：辅助方法和工具类

在 .dna 文件中为每个 .cs 文件添加 `<SourceItem>` 元素。

## 发布和加载说明

生成完毕后，输出以下使用说明：

```
插件已生成在 {outputDir} 目录下：
- {projectName}.xll  （双击或在Excel中拖放加载）
- {projectName}.dna  （同名配置文件，必须在同一目录）
- *.cs                （源代码文件，必须在同一目录）

使用方法：
1. 打开 Excel
2. 将 {projectName}.xll 文件拖放到 Excel 窗口中
3. 按提示允许加载插件
4. 在功能区中查看自定义选项卡
```

## 注意事项

- .dna 文件和 .xll 文件必须同名，且在同一目录下
- 如果 .dna 文件不存在或编译不通过，Excel-DNA 会显示错误提示
- 32位 Excel 需要使用 ExcelDna.xll，64位 Excel 使用 ExcelDna64.xll
- 当前默认使用 64位版本（ExcelDna64.xll）
