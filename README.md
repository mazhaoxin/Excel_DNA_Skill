# README.md

# 项目介绍

用于生成基于`Excel-DNA`的Excel插件的skill。

# 背景介绍

## Excel-DNA
`Excel-DNA`是一个开源工具库，允许开发者使用.NET在Excel中创建高性能自定义函数、宏和插件。它支持多种发布形式，其中最简便易用的方式是将`*.dna`文件与`*.xll`文件一起发布。当在Excel中载入`*.xll`文件时，插件会自动寻找**同名**的`*.dna`文件，编译并加载运行。

## DNA文件
`*.dna`是一个xml格式的文件，主要起到添加引用、定义界面的作用，其格式为：
- Root元素为`DnaLibrary`；
- `DnaLibrary`元素包含一个`Project`元素，用于指定项目的语言和引用，常用C#语言，但**仅能用v4.0的语法**；
- `Project`元素可以包含若干`Reference`元素，用于指定项目的引用，支持Name和Path等。注意：`System`, `System.Collections.Generic`等引用已内置，无需重复引用；
- `Project`元素可以包含若干`SourceItem`元素，用于指定项目的源代码文件；
- `DnaLibrary`元素包含一个`CustomUI`元素，用于指定自定义UI，格式参考Microsoft公司的`Custom UI XML Markup Specification`；
- `CustomUI`元素中引用的事件需要在C#代码中声明；

``` xml
<DnaLibrary Name="MyTest" RuntimeVersion="v4.0">
  <Project Language="C#">
    <Reference Name="System.Drawing" />
    <Reference Name="System.Windows.Forms" />
    <Reference Name="Microsoft.Office.Interop.Excel" />
    <SourceItem Path="MainForm.cs" />
  </Project>
  <CustomUI>
    <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='RibbonLoad'>
      <ribbon>
        <tabs>
          <tab id='CustomTab' label='Custom Tab'>
            <group id='SampleGroup' label='Sample'>
              <button id='Button1' label='Test-1' imageMso='M' size='large' onAction='Button1Click' />
              <button id='Button2' label='Test-2' imageMso='M' size='large' onAction='Button2Click' />
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>
  </CustomUI>
</DnaLibrary>
```

## XLL文件

`*.xll`文件是随`Excel-DNA`发布的已编译过的文件，分成32bit版本（`ExcelDna.xll`）和64bit版本（`ExcelDna64.xll`）。在我的项目中通常使用64bit版本，当前路径为`./ExcelDna-1.9.0/ExcelDna64.xll`。

## CS文件

标准的C#代码文件，但有以下要求
- 必须实现对`ExcelRibbon`类的继承，该类需要添加对`ExcelDna.Integration`和`ExcelDna.Integration.CustomUI`命名空间的引用；
- 必须实现CustomUI中定义的事件，且必须为`public`；
- 必须为C# 4.0的语法，不支持包括`$-string`、`?.`等新版本语法；
- 对Excel文件的操作需要通过`Microsoft.Office.Interop.Excel`进行，`Application`对象必须通过`ExcelDnaUtil.Application`获取，并显式进行数据类型转换；
- `Workbook`, `Worksheet`, `Range`等对象必须进行显式类型转换，因为这些都是COM组件对象；
- CS文件可以有多个，只要在`*.dna`文件中指定即可，不同文件之间定义的类可以跨文件使用。可以根据这个特性合理拆分文件，方便管理和维护。

## 发布
发布时只需要把`ExcelDna64.xll`、`{projectName}.dna`和相关`*.cs`文件复制到同一目录下，并把xll文件重命名为`{projectName}.xll`即可。

## 使用
使用时将`{projectName}.xll`拖放到已打开的Excel窗口中，并按提示允许加载插件即可。如果同名的`*.dna`文件不存在或者编译不通过，会看到相关提示。


# 使用方法

当用户安装过本skill后，可以在`CodeBuddy`对话框中调用skill并描述项目需求，skill会根据需求生成`*.dna`文件和`*.cs`文件，并生成`*.xll`文件。
