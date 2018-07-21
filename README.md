Acrobat DC SDK Documentation


[中文](#-介绍) | [English](#introduction)

![Translate-100%](https://img.shields.io/badge/Translate-100%25-brightgreen.svg)


# 介绍

本章提供了IAC的概念性概述，并介绍了其体系结构和对象层。

使用IAC，外部应用程序可以控制Acrobat或Acrobat Reader。例如，您可以编写启动Acrobat的应用程序，打开特定文件，并设置页面位置和缩放系数。您还可以通过删除页面或添加注释和书签来处理PDF文件。

您的应用程序与Acrobat或Acrobat Reader应用程序之间的通信通过对象和事件进行。

## 关于API对象图层

您可以将Acrobat API视为具有两个使用IAC对象的不同层：

*   Acrobat应用程序（AV）层。通过AV层，您可以控制文档的查看方式。例如，文档对象的视图位于与Acrobat关联的层中。

*   便携式文档（PD）层。PD层提供对文档（例如页面）内的信息的访问。在PD层中，您可以执行PDF文档的基本操作，例如删除，移动或替换页面，以及更改注释属性。您还可以打印PDF页面，选择文本，访问、操作文本以及创建或删除缩略图。

您可以通过使用其PD层对象PDPage或使用其AV层对象AVDoc来控制应用程序的用户界面及其窗口的外观。 PDPage对象有一个名为Draw的方法，它提供了Acrobat的渲染功能。如果需要更精细的控制，可以使用AVDoc对象创建应用程序，该对象具有一个名为OpenInWindow的函数，可以在应用程序窗口中显示文本注释和活动链接。

您还可以将PDF文档视为ActiveX®文档，并通过AcroPDF对象实现方便的PDF浏览器控件。此对象使您能够加载文件，移动文件中的各个页面，以及指定各种显示和打印选项。`OLE对象和方法概要`  章节提供了其用法的详细描述。

### 对象引用语法

Acrobat核心API在C中公开了它的大部分架构，尽管它是为了模拟面向对象的系统而编写的，它有近50个对象。OLE自动化和Apple事件的IAC接口公开了较少数量的对象。这些对象与Acrobat API中的对象紧密相关，可以通过各种编程语言进行访问。

DDE不会围绕对象组织IAC功能，而是使用DDE消息实现Acrobat。

OLE自动化，Apple事件和AppleScript均引用具有不同语法的对象。

*   在OLE中，您在Visual Basic或Visual C＃CreateObject语句或MFC CreateDispatch语句中使用对象名称。

*   在Apple事件中，您在CreateObjSpecifier语句中使用对象的名称。

*   在AppleScript中，您可以在set ... to语句中使用对象名称。

### Acrobat应用程序层中的对象

此表描述Acrobat应用程序（AV）层中的IAC对象。前三个对象是控制用户界面的主要来源。

<table>
<thead>
<tr>
<th>对象</th>
<th>描述</th>
<th>OLE自动化类名</th>
<th>Apple事件类名</th>
</tr>
</thead>
<tbody>
<tr>
<td>AVApp</td>
<td>控制Acrobat的外观。这是顶级对象，代表Acrobat。您可以控制Acrobat的外观，确定是否显示Acrobat窗口，并设置应用程序窗口的大小。您的应用程序可以通过此对象访问菜单栏和工具栏。</td>
<td>AcroExch. 
 App</td>
<td>Application</td>
</tr>
<tr>
<td>AVDoc</td>
<td>表示包含打开的PDF文件的窗口
您的应用程序可以使用此对象
使Acrobat渲染到窗口中，使其与Acrobat窗口非常相似。您还可以使用此对象选择文本，查找文本或打印页面。此对象具有几种访问其他对象的桥接方法。
有关桥方法的详细信息，请参阅`OLE对象和方法概要` 。</td>
<td>AcroExch.
AVDoc</td>
<td>Document</td>
</tr>
<tr>
<td>AVPageView</td>
<td>控制AVDoc窗口的内容。您的应用程序可以滚动，放大或转到下一个，上一个或任意页面。该对象还包含历史堆栈。</td>
<td>AcroExch。
AVPageView</td>
<td>PDF Window</td>
</tr>
<tr>
<td>AVMenu</td>
<td>表示Acrobat中的菜单。您可以计算或删除菜单。每个菜单都有一个与语言无关的名称，用于访问它。</td>
<td>None</td>
<td>Menu</td>
</tr>
<tr>
<td>AVMenuItem</td>
<td>表示菜单中的单个项目。您可以执行或删除菜单项。每个菜单项都有一个与语言无关的名称，用于访问它。</td>
<td>None</td>
<td>Menu item</td>
</tr>
<tr>
<td>AVConversion</td>
<td>表示保存文档的格式。</td>
<td>None</td>
<td>conversion</td>
</tr>
<tr>
<td></td>
<td></td>
<td></td>
<td></td>
</tr>
</tbody>
</table>

### 可移植文档层中的对象

该表描述了便携式文档（PD）层中的IAC对象。

<table>
<thead>
<tr>
<th>目的</th>
<th>描述</th>
<th>OLE自动化类名</th>
<th>Apple事件类名</th>
</tr>
</thead>
<tbody>
<tr>
<td>PDDoc</td>
<td>表示基础PDF文档。使用此对象，您的应用程序可以执行删除和替换页面等操作。您还可以创建和删除缩略图，以及设置和检索文档信息字段。
对于OLE自动化，文档的第一页是0。对于Apple事件，第一页是1。</td>
<td>AcroExch.
PDDoc</td>
<td>Document</td>
</tr>
<tr>
<td>PDPage</td>
<td>表示PDDoc对象的一个​​页面。您可以使用此对象将Acrobat渲染到应用程序的窗口。您还可以访问页面大小和旋转，设置文本区域以及创建和访问注释。
对于OLE自动化，文档的第一页是第0页。对于Apple活动，第一页是第1页。</td>
<td>AcroExch.
PDPage</td>
<td>page</td>
</tr>
<tr>
<td>PDAnnot</td>
<td>处理链接和文本注释。您可以设置和查询注释的物理属性，并可以使用此对象显示链接注释。
Apple事件还有两个相关的对象：PDTextAnnot，一个文本注释，以及一个链接注释PDLinkAnnot。</td>
<td>AcroExch.
PDAnnot</td>
<td>annotation</td>
</tr>
<tr>
<td>PDBookmark</td>
<td>表示PDF文档中的书签。您无法直接创建书签，但如果您知道书签的标题，则可以更改其标题或将其删除。</td>
<td>AcroExch.
PDBookmark</td>
<td>bookmark</td>
</tr>
<tr>
<td>PDTextSelect</td>
<td>使文本显示为已选中。如果选定的文本存在于AVDoc对象中，则应用程序还可以通过此对象访问该区域中的文字。</td>
<td>AcroExch.
PDTextSelect</td>
<td>None</td>
</tr>
<tr>
<td></td>
<td></td>
<td></td>
<td></td>
</tr>
</tbody>
</table>

# OLE对象和方法摘要

OLE自动化由Acrobat API中的一组类提供支持。

下图显示了OLE中使用的对象和方法。箭头表示桥接方法，这些方法可以从不同层的相关对象获取对象。例如，如果你想获得与特定 **AVDoc** 对象相关联的 **PDDoc**，可以使用 **AcroExch.AVDoc** 对象的 **GetPDDoc** 方法。

![OLE](https://help.adobe.com/en_US/acrobat/acrobat_dc_sdk/2015/HTMLHelp/Acro12_MasterBook/IAC_DevApp_OLE_Support/IACMapNew.jpg)

有关完整说明，请参阅`Interapplication Communication API Reference`的OLE自动化部分。

# 使用OLE

本章介绍如何在Adobe Acrobat for Microsoft Windows中使用OLE 2.0支持。Acrobat应用程序是OLE服务器，也响应各种OLE自动化消息。

由于Acrobat提供适当的接口作为OLE服务器，因此您可以将PDF文档嵌入由作为OLE客户端的应用程序创建的文档中，或将它们链接到OLE容器。但是，Acrobat不执行就地激活。

Acrobat支持本章概述的OLE自动化方法，并在`Interapplication Communication API Reference`中完整描述。Acrobat Reader不支持OLE自动化，但AcroPDF对象中提供的PDF浏览器控件除外。

除了对象浏览器之外，Visual Basic或Visual C＃程序员的最佳实用资源是示例项目。这些示例演示了如何使用Acrobat OLE对象，并包含更复杂方法的参数的注释描述。有关更多信息，请参阅`Guide to SDK Samples` 。

本章包含以下信息：

<table>
<thead>
<tr>
<th>话题</th>
<th>描述</th>
</tr>
</thead>
<tbody>
<tr>
<td>Acrobat中的OLE功能</td>
<td>描述了使用OLE进行高级交互式通信时可以做什么。</td>
</tr>
<tr>
<td>开发环境注意事项</td>
<td>描述使用特定开发环境的好处和缺点以及每个环境所需的知识。</td>
</tr>
<tr>
<td>使用Acrobat OLE接口</td>
<td>解释了CAcro和COLEDispatchDriver类的用法。</td>
</tr>
<tr>
<td>使用JSObject接口</td>
<td>解释JSObject接口并提供如何使用它的示例。</td>
</tr>
<tr>
<td>其他开发主题</td>
<td>提供有关OLE自动化的各种信息。</td>
</tr>
<tr>
<td>OLE对象和方法摘要</td>
<td>提供OLE对象和方法的图表以及它们的关联方式。</td>
</tr>
<tr>
<td></td>
<td></td>
</tr>
</tbody>
</table>

有关OLE 2.0和OLE自动化的详细信息，请参阅《OLE Automation Programmer’s Reference》，ISBN 1-55615-851-3，Microsoft Press。您还可以在[http://msdn.microsoft.com上](http://msdn.microsoft.com)找到大量文章。

## Acrobat中的OLE功能

对于OLE自动化，Acrobat提供三种功能：呈现PDF文档，远程控制应用程序以及实现PDF浏览器控件。

### 屏幕渲染

您可以通过两种方式在屏幕上呈现PDF文档：

*   使用类似于Acrobat用户界面的界面。

在此方法中，使用AVDoc对象的OpenInWindowEx方法在你的应用程序窗口中打开PDF文件。窗口有垂直和水平滚动条，窗口周边有按钮，用于设置缩放系数。用户与此类窗口交互发现其操作类似于在Acrobat中的操作。例如，链接处于活动状态，窗口可以在页面上显示任何文本注释。

SDK示例指南中的ActiveView示例显示了如何使用此方法。

*   使用PDPage对象的DrawEx方法。

在此方法中，您提供一个窗口和设备上下文以及缩放系数。Acrobat将当前页面呈现到您的窗口中。应用程序必须管理用户界面中的滚动条和其他项。

SDK示例指南中的StaticView示例显示了如何使用此方法。

### 远程控制Acrobat

您可以通过两种方式远程控制Acrobat：

*   给定输出的接口，您可以编写一个操作PDF文档各个方面的应用程序，例如页面，注释和书签。您的应用程序可能使用AVDoc，PDDoc，PDPage和注释方法，并且可能无法提供任何需要呈现到其应用程序窗口中的视觉反馈。

*   您可以从自己的应用程序启动Acrobat，该应用程序已为用户设置了环境。您的应用程序可以使Acrobat打开文件，设置页面位置和缩放系数，甚至可能选择一些文本。例如，这可以作为帮助系统的一部分。

### PDF浏览器控件

您可以在应用程序中用简化的浏览器控件使用AcroPDF库使显示PDF文档。在这种情况下，PDF文档被视为ActiveX文档，并且该界面在Acrobat Reader中可用。

使用AcroPDF对象的LoadFile方法加载文档。然后，您可以通过浏览器控件实现以下功能：

*   确定要显示的页面

*   选择显示，查看和缩放模式

*   显示书签，缩略图，滚动条和工具栏

*   使用打印页面各种选项

*   高亮显示选择的文本

## 开发环境注意事项

您可以选择与Acrobat集成的环境：Visual Basic，Visual C＃和Visual C ++。

如果可能，请使用Visual Basic或Visual C＃。调用Visual Basic中的CreateObject提供的运行时类型检查允许快速进行应用程序的原型设计，并且在这两种语言中，实现细节都得到了简化。

为了进行比较，请参考以下示例，您可以看到其中有“AcroExch.App”和“Acrobat.CAcroApp” 字符串。第一个是OLE客户端用于创建该类型对象的外部字符串的表单。第二个是包含在开发人员类型库中的表单。

此示例显示了一个Visual Basic子程序，用于查看打开文档的给定页面：

*   使用Visual Basic查看页面

```vb
Private Sub myGoto(ByVal where As Integer)
   Dim app as Object, avdoc as Object, pageview as Object

   Set app = CreateObject("AcroExch.App")
   Set avdoc = app.GetActiveDoc
   Set pageview = avdoc.GetAVPageView
   pageview.Goto(where)
End Sub
```

以下示例在Visual C ++中执行相同操作：

```c++
void goto(int where)
{
   CAcroApp app;
   CAcroAVDoc *avdoc = new CAcroAVDoc;
   CAcroAVPageView pageview;
   COleException e;
   app.CreateDispatch("AcroExch.App");
   avdoc->AttachDispatch(app.GetActiveDoc, TRUE);
   pageview->AttachDispatch(avdoc->GetAVPageView, TRUE);
   pageview->Goto(where);
}
```

下一个示例显示如何使用PDF浏览器控件在Visual Basic中查看页面：

*   在Visual Basic使用AcroPDF浏览器控件

```vb
Friend WithEvents AxAcroPDF1 As AxAcroPDFLib.AxAcroPDF
Me.AxAcroPDF1 = New AxAcroPDFLib.AxAcroPDF

'AxAcroPDF1

Me.AxAcroPDF1.Enabled = True
Me.AxAcroPDF1.Location = New System.Drawing.Point(24, 40)
Me.AxAcroPDF1.Name = "AxAcroPDF1"

Me.AxAcroPDF1.OcxState = CType(
      resources.GetObject("AxAcroPDF1.OcxState"),
      System.Windows.Forms.AxHost.State
)

Me.AxAcroPDF1.Size = New System.Drawing.Size(584, 600)
Me.AxAcroPDF1.TabIndex = 0
AxAcroPDF1.LoadFile("http://www.example.com/example.pdf")
AxAcroPDF1.setCurrentPage(TextBox2.Text)
```

Visual Basic示例更易于阅读，编写和支持，实现细节与Visual C＃类似。
在Visual C ++中，CAcro类隐藏了必须完成的大部分类型检查。在Visual C ++中使用OLE自动化对象需要了COleDispatchDriver类的AttachDispatch和CreateDispatch方法。有关更多信息，请参阅  `使用Acrobat OL接口` 。

> 注意：
> 
>C和C ++程序员使用OLE自动化所需的包含常量值的头文件位于Acrobat DC SDK IAC目录中。Visual Basic和Visual C＃用户不需要这些头文件，尽管引用它们以验证常量定义可能很有用。


### 环境配置
使用Acrobat提供的OLE对象的唯一要求是在您的系统上安装产品，并在项目的项目引用中包含相应的类型库文件Acrobat类型库文件名为Acrobat.tlb。此文件包含在SDK中的  InterAppCommunicationSupport\Headers  件夹中。在项目中包含类型库文件后，可以使用对象浏览器浏览OLE对象。

仅安装ActiveX控件或DLL以启用OLE自动化是不够的。您必须安装完整的Acrobat产品。

如果您是Visual Basic程序员，在项目中包含iac.bas模块（位于headers文件夹中）会很有帮助。该模块定义量变量。

### 必要的C知识
本指南和`Interapplication Communication API Reference`描述了可用的对象和方法。这些文档以及API在计时考虑了C编程，使用API​​编程需要熟悉C概念。

虽然您不需要SDK中提供的头文件，但您可以使用它们来查找文档中引用的各种常量的值，例**AV_DOC_VIEW** 。iac.h 文件中包含大量这些值。

在Visual Basic中使用时，某些方法（如 **OpenInWindowEx** ）最初可能会造成混淆。**OpenInWindowEx** 的 **openflags**参数很**长** 。`Interapplication Communication API Reference`提供的此参数的选项括：

```c
AV_EXTERNAL_VIEW — 打开文档，工具栏可见。
AV_DOC_VIEW — 绘制页面窗格滚动条。
AV_PAGE_VIEW — 仅绘制页面窗格。
```

如果您使用C进行开发，则在编译之前这些字符串将被数字值替换;将这些字符串传递给方法不会引发错误。Visual Basic中编程时，这些字符串对应于iac.bas中定义的常量变量。

在某些情况下，您需要对多个值应用按位OR，并将结果值传递给方法。例如，在 iac.h 中 **PDDocSave** 方法 **ntype** 参数是以下标志按位或:

```c
/* PDSaveFlags — used for PD-level Save 
** All undefined flags should be set to zero. 
** If either PDSaveCollectGarbage or PDSaveCopy are used, PDSaveFull must be used. */
typedef enum { 
   PDSaveIncremental = 0x0000,  /* write changes only */ 
   PDSaveFull = 0x0001,         /* write entire file */ 
   PDSaveCopy = 0x0002,         /* write copy w/o affecting current state */

   PDSaveLinearized = 0x0004,   /* write the file linearized for 
   **       page-served remote (net) access. */

   PDSaveBinaryOK = 0x0010, /* OK to store binary in file */

   PDSaveCollectGarbage = 0x0020  /* perform garbage collection on

   **       unreferenced objects */ 
} PDSaveFlags;
```

例如，如果你想在Visual Basic应用程序内完全保存PDF文件和优化它为Web（线性化），通过 **PDSaveFull** + **PDSaveLinearized**（在 iac.bas 被定义）**ntype** 参数；这相当于 **PDSaveFull** 和 **PDSaveLinearized** 参数的二进制**按位或** 。

在许多情况下，数值在Visual Basic示例代码的注释中已写出来。但是，了解为什么方法以这种方式构造以及如在C中使用它们，对Visual Basic和Visual C＃程序员有用。

## 使用Acrobat OLE接口

本节介绍如何使用CAcro类和COleDispatchDriver类。CAcro类是COleDispatchDriver的子类。

### 关于CAcro类

Acrobat中的OLE 2.0支持几个名称以“ **CAcro** ”开头的类，例如**CAcroApp**和**CAcroPDDoc** 。 SDK的几个文件封装了这些类的定义。

**CAcro**类在Acrobat类型库acrobat.tlb中定义。 Visual Studio中的**OLEView**工具允许您浏览已注册类型库。在Microsoft Visual C ++中使用acrobat.tlb为项目定义OLE自动化。文件acrobat.h和acrobat.cp包含在Acrobat DC SDK中，并为Acrobat自动化服务器实现类型安全的包装器。

> 注意：
> 
>不要修改SDK中的acrobat.tlb，acrobat.h和acrobat.cpp文件;这些定义了Acrobat的OLE自动化界面。

**CAcro**类继承自MFC **COleDispatchDriver**类。理解此类可以更轻松地编写使用 **CAcro** 类及其方法应用程序。

有关**CAcro**类及其方法的详细信息，请参阅`Interapplication Communication API Reference` 。

### 关于COleDispatchDriver类

**COleDispatchDriver**类实现OLE自动化的客户端，提供访问自动化对象所需的大部分代码。它提供了包装函**AttachDispatch** ， **DetachDispatch** 和 **ReleaseDispatch** ，以及便捷函数 **InvokeHelper**， **SetProperty**和 **GetProperty** 。使用Acrobat提供的自动化对象时，您可以使用其中一些方法。其方法使用这些Acrobat中实现的对象。

**COleDispatchDriver**本质上是**IDispatch**的“类包装器”，它是OLE接口，应用程序通过该接口公开方法属性，以便用Visual Basic和Visual C＃编写的其他应用程序可以使用应用程序的功能。这为Acrobat应用程序供OLE支持。

### 使用COleDispatchDriver对象和方法

本节讨论如何使用acrobat.cpp导出的类，并展示何时调用 **CreateDispatch** 和 **AttachDispatch** 方法。
以下是acrobat.h中的一段代码，它声明了 **CAcroHiliteList** 类。**CAcroHiliteList" 是 **COleDispatchDriver** 类的子类，这意味着它共享 **COleDispatchDriver** 的所有实例变量。

其中一个变量是 **m_lpDispatch** ，它包含该对象的 **LPDISPATCH** 。**LPDISPATCH**是指 **IDispatch** 的 **长** 指针，可以将其视为表示调度连接的不透明数据类型。 **m_lpDispatch** 可用于需 **LPDISPATCH** 参数的函数。

*   CAcroHiliteList类声明

```c++
class CAcroHiliteList : public COleDispatchDriver
{
public:
   CAcroHiliteList() {}        // Calls COleDispatchDriver default constructor
   CAcroHiliteList(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
   CAcroHiliteList(const CAcroHiliteList& dispatchSrc) :
      COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
   bool Add(short nOffset, short nLength);
};
```

以下是acrobat.cpp中Add方法的相关实现部分：

```c++
bool CAcroHiliteList::Add(short nOffset, short nLength)
{
   bool result;
   static BYTE parms[] =
      VTS_I2 VTS_I2;
   InvokeHelper(0x1, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
      nOffset, nLength);
   return result;
}
```

调用Add方法时，例如`Using the COleDispatchDriver class` 下示例中的代码，

```c++
hilite->Add(0, 10);
```

调用 **InvokeHelper** 函数。此 **COleDispatchDriver** 方法采用可变数量的参数。它最终调用了 **CAcroHiliteList** 对象的Add方法的Acrobat实现。这发生在虚拟OLE“wires”上，并处理所有OLE细节。最终结果是将页面范围添加到 **CAcroHiliteList** 对象。

以下是从ActiveView示例改编的方法的实现：

*   使用COleDispatchDriver类

``` c++
// This code demonstrates how to highlight words with
// either a word or page highlight list

void CActiveViewDoc::OnToolsHilitewords()
{
   CAcroAVPageView pageView;
   CAcroPDPage page;
   CAcroPDTextSelect* textSelect = new CAcroPDTextSelect;
   CAcroHiliteList* hilite = new CAcroHiliteList;
   char buf[255];
   long selectionSize;

   if ((BOOL) GetCurrentPageNum() > PDBeforeFirstPage) {
      // Obtain the AVPageView
      pageView.AttachDispatch(m_pAcroAVDoc->GetAVPageView(),TRUE);

      // Create the Hilite list object
      hilite->CreateDispatch("AcroExch.HiliteList");
      if (hilite) {

   // Add the first 10 words or characters of that page to the highlight list
         hilite->Add(0,10);
         page.AttachDispatch(pageView.GetPage(), TRUE);

         // Create text selection for either page or word highlight list
         textSelect->AttachDispatch(page.CreateWordHilite(hilite->m_lpDispatch));
         m_pAcroAVDoc->SetTextSelection(textSelect->m_lpDispatch);
         m_pAcroAVDoc->ShowTextSelect();

         // Extract the number of words and the first word of text selection
         selectionSize = textSelect->GetNumText();
         if (selectionSize)
            sprintf (buf, "# of words in text selection: %ld\n1st word in text
               selection = '%s'", selectionSize, textSelect->GetText(0));
         else
            sprintf (buf, "Failed to create text selection.");

         AfxMessageBox(buf);
      }
   }

   delete textSelect;
   delete hilite;
}
```

在前面的示例中，具有前缀 **CAcro** 的对象都是 **CAcro** 类对象 - 它们也是 **COleDispatchDriver** 对象 - 因为所有Acrobat **CAcro** 类都是 **COleDispatchDriver** 的子类。

_实例化类不足以使用它_ 。在使用对象之前，必须使用 **COleDispatchDriver** 类的 **Dispatch** 的一个方法将对象附加到相应的Acrobat对象。这些函数还初始化对象的 **m_lpDispatch** 实例变量。

上一个示例中的代码显示了如何附加已存在的 **IDispatch** ：

```c++
CAcroAVPageView pageView;
// Obtain the AVPageView 
pageView.AttachDispatch(m_pAcroAVDoc->GetAVPageView(), TRUE);
```

**CAcroAVDoc** 类的 **GetAVPageView** 方法返回一个 **LPDISPATCH** ，这是 **AttachDispatch** 方法期望的第一个参数。作为第二个参数传递的 **BOOL** 指示当对象超出范围时是否应释放 **IDispatch** ，并且通常为 **TRUE** 。通常，从 **GetAVPageView** 等方法返回 **LPDISPATCH** 时 ，使用 **AttachDispatch** 将其附加到对象。

上一个示例中的以下代码使用 **CreateDispatch** 方法：

```c++
CAcroHiliteList *hilite = new CAcroHiliteList;
hilite->CreateDispatch("AcroExch.HiliteList");
hilite->Add(0, 10);
```

在这种情况下， **CreateDispatch** 方法都会创建 **IDispatch** 对象并将其附加到对象。这段代码工作正常；但是以下代码将失败：

```c++
CAcroHiliteList *hilite = new CAcroHiliteList;
hilite->Add(0, 10);
```

此错误类似于使用未初始化的变量。在 **IDispatch** 对象附加到 **COleDispatchDriver** 对象之前，它无效。

**CreateDispatch** 采用字符串参数，例如“ **AcroExch.HiliteList** “，代表一个类。以下代码不正确：

```c++
CAcroPDDoc doc = new CAcroPDDoc;
doc.CreateDispatch("AcroExch.Create");
```

这会失败，因为Acrobat不会响应这样的参数。参数应为“ **AcroExch.PDDoc** “。

CreateDispatch的有效字符串如下：

| 类 | 字符串 |
| -- | -- |
| CAcroPoint | "AcroExch.Point" |
| CAcroRect | "AcroExch.Rect" |
| CAcroTime | "AcroExch.Time" |
| CAcroApp | "AcroExch.App" |
| CAcroPDDoc | "AcroExch.PDDoc" |
| CAcroAVDoc | "AcroExch.AVDoc" |
| CAcroHiliteList | "AcroExch.HiliteList" |
| CAcroPDBookmark | "AcroExch.PDBookmark" |
| CAcroMatrix | "AcroExch.Matrix" |
| AcroPDF | "AxAcroPDFLib.AxAcroPDF" |
|  |  |

再次参考上一个示例中的代码：

```c++
CAcroPDPage page;
page.AttachDispatch(pageView.GetPage(), TRUE);
```

**PDPage** 对象是必需的，因为此代码的目的是突出显示当前页面上的文字。由于它是一个 **CAcro** 变量，因此必须在使用其方法之前附加到OLE对象。**CreateDispatch** 不能用于创建PDPage对象，因为“ **AcroExch.PDPage** “不是 **CreateDispatch** 的有效字符串。但是， **AVPageView** 方法 **GetPage** 返回 **PDPage** 对象的 **LPDISPATCH** 指针。这作为页面对象的 **AttachDispatch** 方法的第一个参数传递。**TRUE** 参数表示当对象超出范围时将自动释放该对象。

```c++
CAcroPDTextSelect* textSelect = new CAcroPDTextSelect;
textSelect->AttachDispatch
   (page.CreateWordHilite(hilite->m_lpDispatch));
m_pAcroAVDoc->SetTextSelection (textSelect->m_lpDispatch);
m_pAcroAVDoc->ShowTextSelect();
```

此代码执行以下步骤：

1.  声明文本选择对象textSelect。
2.  调用 **CAcroPDPage** 方法 **CreateWordHilite** ，它返回PDTextSelect的LPDISPATCH。**CreateWordHilite**采用表示**CAcroHilite**列表的LPDISPATCH参数。该 **HILITE** 变量已经包含了 **CAcroHiliteList** 对象，它的实例变量 **m_lpDispatch** 包含该对象的指针LPDISPATCH。
3.  调用 **CAcroAVDoc** 对象的SetTextSelection方法以选择当前页面上的前十个文字。
4.  调用 **AcroAVDoc** 的ShowTextSelect方法以在屏幕上进行显示更新。

## 使用JSObject接口

Acrobat提供了一组丰富的JavaScript编程接口，可以在Acrobat环境中使用。它还提供了 **JSObject** 接口，允许外部客户端从Visual Basic等环境访问相同的功能。

准确地说， **JSObject** 是OLE自动化客户端（如Visual Basic应用程序）与Acrobat提供的JavaScript功能之间的解释层。从开发人员的角度来看，在Visual Basic环境中编写 **JSObject** 类似于使用Acrobat控制台在JavaScript中编程。

本节介绍如何在Visual Basic编程环境中使用JavaScript扩展Acrobat。它提供了一组示例来说明关键概念。

只要有可能，您应该使用 **AcroExch..PDDoc** 对象中提供的 **JSObject**接口来利用这些功能**。要获取接口，请调用对象的 **GetJSObject** 方法。

### 添加对Acrobat类型库的引用

此过程添加对Acrobat类型库的引用，以便您可以在Visual Basic中访问Acrobat自动化API，包括JSObject。在使用JSObject接口之前执行此操作，如下面的示例所示。

要添加对Acrobat类型库的引用：

1.  安装Acrobat和Visual Basic。
2.  从Windows应用程序模板创建一个新的Visual Basic项目。这提供了一个空白表单和项目工作区。
3.  选择 **Project** > **Add Reference** ，然后单击 **COM** 选项卡。
4.  从可用引用列表中，选择“ **Adobe Acrobat 8​​.0类型库”** ，然后单击“ **OK** 。

### 创建一个简单的应用程序

此示例提供显示“Hello，Acrobat！”的最小代码。在Acrobat JavaScript控制台中。

设置并运行“Hello，Acrobat！”示例：

1.  单击“view” > “code”，打开默认表单的源代码窗口。
2.  从该窗口左上角的选择框中选择（Form1 Events）。

右上角的选择框现在显示Form1对象可用的所有功能。

3.  从功能选择框中选择加载函数。这会创建一个空函数存根。首次显示Form1时，将调用Form1 Load函数，因此这是添加初始化代码的好地方。
4.  添加以下代码以在子例程之前定义一些全局变量。

```vb
Dim gApp As Acrobat.CAcroApp
Dim gPDDoc As Acrobat.CAcroPDDoc
Dim jso As Object
```

5.  将以下代码添加到私有Form1_Load子程序。

```vb
gApp = CreateObject("AcroExch.App")
gPDDoc = CreateObject("AcroExch.PDDoc")

If gPDDoc.Open("c:\example.pdf") Then
    jso = gPDDoc.GetJSObject
    jso.console.Show
    jso.console.Clear
    jso.console.println ("Hello, Acrobat!")
    gApp.Show
End If
```

6.  在C：驱动器的根级别创建名为example.pdf的文件。
7.  保存并运行该项目。

运行应用程序时，将启动Acrobat，显示Form1，并打开JavaScript Debugger窗口，显示“Hello，Acrobat！”。

*   显示“Hello，Acrobat！”在JavaScript控制台中

```vb
Dim gApp As Acrobat.CAcroApp
Dim gPDDoc As Acrobat.CAcroPDDoc
Dim jso As Object

Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs)
      Handles Me.Load
   gApp = CreateObject("AcroExch.App")
   gPDDoc = CreateObject("AcroExch.PDDoc")
   If gPDDoc.Open("c:\example.pdf") Then
      jso = gPDDoc.GetJSObject
      jso.console.Show
      jso.console.Clear
      jso.console.println ("Hello, Acrobat!")
      gApp.Show
   End If
End Sub
```

Visual Basic程序使用 **CreateObject** 调用附加到Acrobat自动化接口，然后使用App对象的 **Show** 命令显示主窗口。

在研究代码后，您可能会遇到一些问题。例如，为什么jso声明为Object，而 **gApp** 和 **gPDDoc** 声明为Acrobat类型库中的类型？ **JSObject** 有真正的类型吗？

答案是否定的， **JSObject** 没有出现在类型库中，除了在 **CAcroPDDoc.GetJSObject** 调用的上下文中。用于通过JSObject导出JavaScript功能的COM接口称为IDispatch接口，在Visual Basic中，它通常简称为“对象”类型。这意味着程序员可以使用的方法没有特别明确的定义。例如，如果您将调用替换为

```vb
jso.console.clear
```

及

```vb
jso.ThisCantPossiblyCompileCanIt("Yes it can!")
```

编译器编译代码，但在运行时失败。Visual Basic没有 **JSObject** 的类型信息，因此Visual Basic不知道特定调用在运行时是否在语法上有效，并且将编译对 **JSObject** 的所有函数调用。因此，您必须依赖文档来了解通过 **JSObject** 接口可用的功能。有关详细信息，请参阅`JavaScript for Acrobat API Reference` 。

您可能也想知道为什么在创建 **JSObject** 之前必须打开PDDoc。运行程序显示屏幕上没有文档，建议在没有 **PDDoc的** 情况下使用JavaScript控制台。但是 **JSObject** 旨在与特定文档紧密 **协作** ，因为大多数可用功能在文档级别运行。JavaScript中有一些应用程序级功能（在 **JSObject** 中），但它们是次要的。实际上， **JSObject** 始终与特定文档相关联。

处理大量文档时，必须构造代码，以便为每个文档获取新的 **JSObject** ，而不是创建单个 **JSObject** 来处理每个文档。

### 使用注释

此示例使用JSObject接口打开PDF文件，向其添加预定义注释，并将文件保存回磁盘。

要设置并运行注释示例：

1.  创建一个新的Visual Basic项目并将Adobe Acrobat类型库添加到项目中。
2.  从工具箱中，将 **OpenFileDialog** 控件拖到窗体中。
3.  将 **Button** 拖到表单中。

    ![IMG](https://help.adobe.com/en_US/acrobat/acrobat_dc_sdk/2015/HTMLHelp/Acro12_MasterBook/IAC_DevApp_OLE_Support/addingbutton.gif)

1.  选择 **View** > **Code**并设置以下源代码：

*   添加注释

```vb
Dim gApp As Acrobat.CAcroApp

Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As
      System.EventArgs) Handles MyBase.Load
   gApp = CreateObject("AcroExch.App")
End Sub

Private Sub Form1_Closed(Cancel As Integer)
   If Not gApp Is Nothing Then
      gApp.Exit
   End If
   gApp = Nothing
End Sub

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As
      System.EventArgs) Handles Button1.Click
   Dim pdDoc As Acrobat.CAcroPDDoc
   Dim page As Acrobat.CAcroPDPage
   Dim jso As Object
   Dim path As String
   Dim point(1) As Integer
   Dim popupRect(3) As Integer
   Dim pageRect As Object
   Dim annot As Object
   Dim props As Object
   
   OpenFileDialog1.ShowDialog()
   path = OpenFileDialog1.FileName
   
   pdDoc = CreateObject("AcroExch.PDDoc")
   If pdDoc.Open(path) Then
      jso = pdDoc.GetJSObject
      If Not jso Is Nothing Then
      
         ' Get size for page 0 and set up arrays
         page = pdDoc.AcquirePage(0)
         pageRect = page.GetSize
         point(0) = 0
         point(1) = pageRect.y
         popupRect(0) = 0
         popupRect(1) = pageRect.y - 100
         popupRect(2) = 200
         popupRect(3) = pageRect.y
         
         ' Create a new text annot
         annot = jso.AddAnnot
         props = annot.getProps
         props.Type = "Text"
         annot.setProps props
         
         ' Fill in a few fields
         props = annot.getProps
         props.page = 0
         props.point = point
         props.popupRect = popupRect
         props.author = "John Doe"
         props.noteIcon = "Comment"
         props.strokeColor = jso.Color.red
         props.Contents = "I added this comment from Visual Basic!"
         annot.setProps props
      End If
      pdDoc.Close
      MsgBox "Annotation added to " & path
   Else
      MsgBox "Failed to open " & path
   End If
   
   pdDoc = Nothing
End Sub
```

5.  保存并运行该应用程序。

**Form_Load** 和 **Form_Closed** 例程中的代码初始化并关闭Acrobat自动化接口。例程中命令按钮的单击会发生更有趣的工作。第一行声明局部变量并显示Windows Open对话框，允许用户选择要注释的文件。然后代码打开PDF文件的 **PDDoc** 对象并获取该文档的 **JSObject** 接口。

一些标准的Acrobat自动化方法用于确定文档中第一页的大小。这些数字对于实现正确的布局至关重要，因为PDF坐标系位于页面的左下角，但注释将锚定在页面的左上角。

“ **Create a new text annot**”注释后面的行就是这样做的，但这段代码带有额外的解释。

首先， **addAnnot** 看起来好像是 **JSObject** 的方法，但JavaScript引用显示该方法与 **doc** 对象相关联。您可能希望语法为 **jso.doc.addAnnot** 。因为 **jso** 是 **Doc** 对象，所以 **jso.addAnnot** 是正确的。**Doc** 对象中的所有属性和方法都以这种方式使用。

其次，观察 **annot.getProps** 和 **annot.setProps** 的使用。Annot对象使用单独的属性对象实现，这意味着您无法直接设置属性。例如，您无法执行以下操作：

```vb
annot = jso.AddAnnot
annot.Type = "Text"
annot.page = 0
...
```

相反，您必须使用 **annot.getProps** 获取 **Annot** 的属性对象，并使用该对象进行读取或写入访问。要将更改保存回原始 **Annot** ，请使用修改后的属性对象调用 **annot.setProps** 。

第三，注意使用 **JSObject** 的color属性。此对象定义了几种简单的颜色，如红色，绿色和蓝色。在处理颜色时，您可能需要比通过此对象可用的颜色范围更大的颜色。此外，每次调用 **JSObject** 都会产生性能 **损失** 。要更有效地设置颜色，可以使用以下代码，它将annot的 **strokeColor** 直接设置为红色，绕过颜色对象。

```vb
dim color(0 to 3) as Variant
color(0) = "RGB"
color(1) = 1#
color(2) = 0#
color(3) = 0#
annot.strokeColor = color
```

您可以在需要颜色数组作为 **JSObject** 例程的参数的任何地方使用此技术。该示例将颜色空间设置为RGB，并指定红色，绿色和蓝色的浮点值，范围从0到1。请注意在颜色值后面使用＃字符。这些是必需的，因为它们告诉Visual Basic数组元素应该设置为浮点值而不是整数。将数组声明为包含变量也很重要，因为它包含字符串和浮点值。其他颜色空间（“T”，“G”，“CMYK”）对数组长度有不同的要求。有关更多信息，请参阅`JavaScript for Acrobat API Reference`的 **Color** 对象。

> 注意：
> 
>如果您希望用户能够编辑注释，请将JavaScript属性Collab.showAnnotsToolsWhenNoCollab设置为true。

### 文档拼写检查

Acrobat包含一个插件，可以扫描文档以查找拼写错误。该插件还提供了可以使用 **JSObject** 访问的JavaScript方法。在此示例中，您将从示例`Adding an annotation`的源代码开始，并进行以下更改：

*   将列表视图控件添加到主窗体。保留控件的默认名称ListView1。
*   使用以下内容替换现有Command1_Click例程中的代码：
*   拼写检查文档

```vb
Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As
      System.EventArgs) Handles Button1.Click
   Dim pdDoc As Acrobat.CAcroPDDoc
   Dim jso As Object
   Dim path As String
   Dim count As Integer
   Dim i As Integer, j As Integer
   Dim word As Variant
   Dim result As Variant
   Dim foundErr As Boolean
   
   OpenFileDialog1.ShowDialog()
   path = OpenFileDialog1.FileName
   foundErr = False
   pdDoc = CreateObject("AcroExch.PDDoc")
   
   If pdDoc.Open(path) Then
      jso = pdDoc.GetJSObject
      If Not jso Is Nothing Then
         count = jso.getPageNumWords(0)
         For i = 0 To count - 1
            word = jso.getPageNthWord(0, i)
            If VarType(word) = vbString Then
               result = jso.spell.checkWord(word)
               If IsArray(result) Then
                  foundErr = True
                  ListView1.Items.Add (word & " is misspelled.")
                  ListView1.Items.Add ("Suggestions:")
                  For j = LBound(result) To UBound(result)
                     ListView1.Items.Add (result(j))
                  Next j
                  ListView1.Items.Add ("")
               End If
            End If
         Next i
         jso = Nothing
         pdDoc.Close
         
         If Not foundErr Then
            ListView1.Items.Add ("No spelling errors found in " & path)
         End If
      End If
   Else
      MsgBox "Failed to open " & path
   End If
   
   pdDoc = Nothing
End Sub
```

在此示例中，请注意使用Spell对象的 **check** 方法。如`JavaScript for Acrobat API Reference` 中描述，此方法将单词作为输入，如果在单词中找到单词，则返回null对象；如果找不到单词，则返回建议单词数组。

存储 **JSObject** 方法调用的返回值时，最安全的方法是使用Variant。您可以使用 **IsArray** 函数来确定Variant是否是一个数组，并编写代码来相应地处理该情况。在这个简单的示例中，如果程序找到建议单词的数组，它会将它们转储到List View控件。

### 将JavaScript转换为JSObject的提示

涵盖JSObject可用的每种方法都超出了本文档的范围。但是， `JavaScript for Acrobat API Reference`的`JavaScript for Acrobat API Reference`详细介绍了主题，并且可以通过记住一些基本事实从参考中推断出很多内容：

*   参考中的大多数对象和方法都可以在Visual Basic中使用，但不是全部。特别是，无法在Visual Basic中创建任何需要构造新运算符的JavaScript对象。这包括Report对象。
*   Annots对象的不寻常之处在于它需要JSObject使用getProps和setProps方法将其作为单独的对象来设置和获取属性。
*   如果您不确定用于声明变量的类型，请将其声明为Variant。这为Visual Basic提供了更多的类型转换灵活性，并有助于防止运行时错误。
*   JSObject无法向JavaScript添加新属性，方法或对象。由于此限制， **global.setPersistent**属性没有意义。
*   JSObject不区分大小写。Visual Basic通常会将标识符的前导字符大写，并阻止您更改其大小写。不要担心这一点，因为JSObject在将标识符与其JavaScript等效项匹配时忽略了大小写。
*   JSObject始终将值返回为Variants。这包括属性获取以及方法调用的返回值。当预期返回null值时，将使用空Variant。当JSObject返回一个数组时，数组中的每个元素都是一个Variant。若要确定Variant的实际数据类型，请使用Visual Basic for Applications（VBA）库的Information模块中的实用程序函数 **IsArray** ， **IsNumeric** ， **IsEmpty** ， **IsObject**和VarType。
*   JSObject可以处理Visual Basic的大多数类型，用于设置属性和方法调用的参数，包括Variant，Array，Boolean，String，Date，Double，Long，Integer和Byte。JSObject可以接受Object参数，但仅当Object是对JSObject的属性get或方法调用的结果时。JSObject无法接受Error和Currency类型的值。

## 其他开发主题

本节包含与开发OLE应用程序相关的各种主题。

### 同步消息传递

Acrobat OLE自动化实现基于同步消息传递方案。当应用程序向Acrobat发送请求时，应用程序处理该请求并将控制权返回给应用程序。只有这样，应用程序才能向Acrobat发送另一条消息。如果您的应用程序发送一条消息，紧接着另一条消息，则可能无法正确接收第二条消息：它不会生成服务器繁忙错误，而是失败，没有错误消息。

例如，使用 **AVDoc.OpenInWindowEx** 方法可以实现这一点 ，其中交互了大量关于绘图位置和鼠标点击的信息，以及 **PDPage.DrawEx** 方法的用法在特别复杂的页面上。使用 **DrawEx** 方法，在生成WM_PAINT消息时会出现问题。如果页面很复杂且环境是多线程的，则应用程序可能无法在应用程序生成另一个WM_PAINT消息之前完成绘制页面。由于应用程序是单线程的，因此多线程应用程序必须适当地处理这种情况。

### MDI应用程序

假设您创建了一个多文档界面（MDI）应用程序，该应用程序创建了一个静态窗口，使用 **OpenInWindowEx** 调用将Acrobat显示在该窗口中，此窗口基于 **CFormView** OLE类。如果另一个窗口放在该窗口的顶部并随后被删除，则Acrobat窗口不会正确重新绘制。

要解决此问题，请将子选夹样式指定给对话框模板（ **CFormView** 所基于的模板）。否则，该对话框将删除所有子窗口的背景，包括包含PDF文件的窗口，该窗口将擦除以前覆盖的PDF窗口部分。

### 子窗口中的事件处理

使用 **OpenInWindowEx** 打开PDF文件时，Acrobat会在其上创建子窗口。这允许应用程序直接接收此窗口的事件。但是，应用程序还必须处理以下事件： **resize** ， **key up** 和 **key down** 。

ActiveView示例中的以下示例显示了如何处理resize事件：

*   处理调整大小事件

```c++
void CActiveViewVw::OnSize(UINT nType, int cx, int cy)
{
   CWnd* pWndChild = GetWindow(GW_CHILD);
   if (!pWndChild)
      return;
   CRect rect;
   GetClientRect(&rect);
   pWndChild->
      SetWindowPos(NULL,0,0,rect.Width,rect.Height,
               SWP_NOZORDER | SWP_NOMOVE);

   CView::OnSize(nType, cx, cy);
}
```

将消息发送到子窗口后，它也会进行调整大小。这导致两个窗口都被调整大小，这是期望的效果。

### 确定Acrobat应用程序是否正在运行

将Windows FindWindow方法与Acrobat类名一起使用。您可以使用Microsoft Spy ++实用程序来确定应用程序版本的类名。

### 退出应用程序

当用户使用OLE自动化退出应用程序时，Acrobat本身或显示PDF文档的Web浏览器可能会受到影响：

*   如果Acrobat中未打开PDF文档，则应用程序将退出。
*   如果Web浏览器正在显示PDF文档，则显示为空白。用户可以刷新页面以重新显示它。

## 使用DDE

本章介绍Microsoft Windows下Acrobat中的DDE支持。虽然支持DDE，但应尽可能使用OLE自动化而不是DDE，因为DDE不是COM技术。

有关与DDE消息关联的参数的完整描述，请参阅`Interapplication Communication API Reference`的DDE部分。

对于所有DDE消息，服务名称为acroview，事务类型为XTYPE_EXECUTE，主题名称为control。数据是要执行的命令，括在方括号内。DdeClientTransaction调用中的item参数为NULL。

以下示例设置DDE消息：

*   设置DDE消息

```vb
DDE_SERVERNAME = "acroview";
DDE_TOPICNAME = "control";
DDE_ITEMNAME = "[AppHide()]";
```

DDE消息中的方括号字符是必需的。DDE消息区分大小写，必须完全按照描述使用。

为了能够在文档上使用DDE消息，必须首先使用DocOpen DDE消息打开文档。您无法使用DDE消息关闭用户手动打开的文档。

您可以对路径名使用NULL，在这种情况下，DDE消息在前端文档上运行。

如果一次发送多个命令，则按顺序执行命令，结果将作为单个操作显示给用户。例如，您可以使用此功能将文档打开到特定页面和缩放级别。

页码从零开始：文档中的第一页是第0页。仅当参数包含空格时才需要引号。

文档操作方法（例如用于删除页面或滚动的方法）仅适用于已打开的文档。

## 使用Apple Events

您可以使用多个对象和事件来为Mac OS开发Acrobat应用程序。支持Apple事件注册表中的某些对象和事件，以及特定于Acrobat的对象和事件。Acrobat支持以下类别的Apple事件：

<table>
<thead>
<tr>
<th>类别</th>
<th>描述</th>
</tr>
</thead>
<tbody>
<tr>
<td>必需的活动</td>
<td>Finder发送给所有应用程序的事件。</td>
</tr>
<tr>
<td>核心事件</td>
<td>各种应用程序通用的事件，但并非普遍适用于所有应用程序。</td>
</tr>
<tr>
<td>特定于Acrobat的事件</td>
<td>特定于Acrobat的事件。</td>
</tr>
<tr>
<td>其他Apple事件</td>
<td>不属于上述类别之一的事件。</td>
</tr>
<tr>
<td></td>
<td></td>
</tr>
</tbody>
</table>

在为Mac OS编程时，尽可能将AppleScript与Acrobat一起使用。对于AppleScript无法使用的Apple事件，请使用C或其他编程语言处理它们。有关参数的完整说明，请参阅“ `Interapplication Communication API Reference` 。

有关Acrobat Search插件支持的Apple事件的信息，请参阅`Acrobat and PDF Library API Reference` 。有关支持其他Apple事件的其他插件的信息，请参阅`Overview` 。

有关Apple事件和脚本的详细信息，请参阅_Inside Macintosh：Interapplication Communication_ ，ISBN 0-201-62200-9，Addison-Wesley。本文档的内容目前可从[http://developer.apple.com/documentation/mac/IAC/IAC-2.html](http://developer.apple.com/documentation/mac/IAC/IAC-2.html) 获取。

有关AppleScript语言的详细信息，请参阅AppleScript语言指南，ISBN 0-201-40735-3，Addison-Wesley。该文档的内容目前可从[http://developer.apple.com/documentation/AppleScript/Conceptual/AppleScriptLangGuide/](http://developer.apple.com/documentation/AppleScript/Conceptual/AppleScriptLangGuide/) 获取。

有关核心和所需Apple事件的更多信息，请参阅适用于Mac OS的Apple事件注册表。此文件位于AppleScript 1.3.4 SDK中，该SDK目前可从[http://developer.apple.com/sdk/](http://developer.apple.com/sdk/) 获取。


# Introduction

This chapter provides a conceptual overview to IAC and introduces its architecture and object layers.

With IAC, an external application can control Acrobat or Acrobat Reader. For example, you can write an application that launches Acrobat, opens a specific file, and sets the page location and zoom factor. You can also manipulate PDF files by, for example, deleting pages or adding annotations and bookmarks.

Communication between your application and the Acrobat or Acrobat Reader application occurs through objects and events.

## About the API object layers

You can think of the Acrobat API as having two distinct layers that use IAC objects:

- The Acrobat application (AV) layer. The AV layer enables you to control how the document is viewed. For example, the view of a document object resides in the layer associated with Acrobat.

- The portable document (PD) layer. The PD layer provides access to the information within a document, such as a page. From the PD layer you can perform basic manipulations of PDF documents, such as deleting, moving, or replacing pages, as well as changing annotation attributes. You can also print PDF pages, select text, access manipulated text, and create or delete thumbnails.

You can control the application’s user interface and the appearance of its window by either using its PD layer object, PDPage, or by using its AV layer object, AVDoc. The PDPage object has a method called Draw that exposes the rendering capabilities of Acrobat. If you need finer control, you can create your application with the AVDoc object, which has a function called OpenInWindow that can display text annotations and active links in your application’s window.

You can also treat a PDF document as an ActiveX® document and implement convenient PDF browser controls through the AcroPDF object. This object provides you with the ability to load a file, move to various pages within a file, and specify various display and print options. A detailed description of its usage is provided in ``Summary of OLE objects and methods``.

### Object reference syntax

The Acrobat core API exposes most of its architecture in C, although it is written to simulate an object-oriented system with nearly fifty objects. The IAC interface for OLE automation and Apple events exposes a smaller number of objects. These objects closely map to those in the Acrobat API and can be accessed through various programming languages.

DDE does not organize IAC capabilities around objects, but instead uses DDE messages to Acrobat.

OLE automation, Apple events, and AppleScript each refer to the objects with a different syntax.

- In OLE, you use the object name in either a Visual Basic or Visual C# CreateObject statement or in an MFC CreateDispatch statement.

- In Apple events, you use the name of the object in a CreateObjSpecifier statement.

- In AppleScript, you use the object name in a set ... to statement.

### Objects in the Acrobat application layer

This table describes the IAC objects in the Acrobat application (AV) layer. The first three objects are the primary source for controlling the user interface.

| Object | Description | OLE automation class name | Apple event class name |
| -- | -- | -- | -- |
| AVApp | Controls the appearance of Acrobat. This is the top-level object, representing Acrobat. You can control the appearance of Acrobat, determine whether an Acrobat window appears, and set the size of the application window. Your application has access to the menu bar and the toolbar through this object. | AcroExch. <br> App | Application |
| AVDoc | Represents a window containing an open PDF<br>file. Your application can use this object to<br>cause Acrobat to render into a window so that it closely resembles the Acrobat window. You can also use this object to select text, find text, or print pages. This object has several bridge methods to access other objects.<br>For more information on bridge methods, see ``Summary of OLE objects and methods``. | AcroExch.<br>AVDoc | Document |
| AVPageView | Controls the contents of the AVDoc window. Your application can scroll, magnify, or go to the next, previous, or any arbitrary page. This object also holds the history stack. | AcroExch.<br>AVPageView | PDF Window |
| AVMenu | Represents a menu in Acrobat. You can count or remove menus. Each menu has a language-independent name used to access it. | None | Menu |
| AVMenuItem | Represents a single item in a menu. You can execute or remove menu items. Every menu item has a language-independent name used to access it. | None | Menu item |
| AVConversion | Represents the format in which to save the document. | None | conversion |
| | | |

### Objects in the portable document layer

This table describes the IAC objects in the portable document (PD) layer.

| Object | Description | OLE automation class name | Apple event class name |
| -- | -- | -- | -- |
| PDDoc | Represents the underlying PDF document. Using this object, your application can perform operations such as deleting and replacing pages. You can also create and delete thumbnails, and set and retrieve document information fields.<br>For OLE automation, the first page of a document is page 0. For Apple events, the first page is page 1. | AcroExch.<br>PDDoc | Document |
| PDPage | Represents one page of a PDDoc object. You can use this object to render Acrobat to your application’s window. You can also access page size and rotation, set up text regions, and create and access annotations.<br>For OLE automation, the first page of a document is page 0. For Apple events, the first page is page 1. | AcroExch.<br>PDPage | page |
| PDAnnot | Manipulates link and text annotations. You can set and query the physical attributes of an annotation and you can perform a link annotation with this object.<br>Apple events have two additional, related objects: PDTextAnnot, a text annotation, and PDLinkAnnot, a link annotation. | AcroExch.<br>PDAnnot | annotation |
| PDBookmark | Represents bookmarks in the PDF document. You cannot directly create a bookmark, but if you know a bookmark’s title, you can change its title or delete it. | AcroExch.<br>PDBookmark | bookmark |
| PDTextSelect | Causes text to appear selected. If selected text exists within an AVDoc object, your application can also access the words in that region through this object. | AcroExch.<br>PDTextSelect | None |
| | | |

# Summary of OLE objects and methods
OLE automation support is provided by a set of classes in the Acrobat API.

The following diagram shows the objects and methods that are used in OLE. The arrows indicate bridge methods, which are methods that can get an object from a related object of a different layer. For example, if you want to get the **PDDoc** associated with a particular **AVDoc** object, you can use the **GetPDDoc** method in the **AcroExch.AVDoc** object.

![ole](https://help.adobe.com/en_US/acrobat/acrobat_dc_sdk/2015/HTMLHelp/Acro12_MasterBook/IAC_DevApp_OLE_Support/IACMapNew.jpg)

For complete descriptions, see the OLE automation sections of the ``Interapplication Communication API Reference``.

# Using OLE

This chapter describes how you can use OLE 2.0 support in Adobe Acrobat for Microsoft Windows. Acrobat applications are OLE servers and also respond to a variety of OLE automation messages.

Since Acrobat provides the appropriate interfaces to be an OLE server, you can embed PDF documents into documents created by an application that is an OLE client, or link them to OLE containers. However, Acrobat does not perform in-place activation.

Acrobat supports the OLE automation methods that are summarized in this chapter and described fully in the ``Interapplication Communication API Reference``. Acrobat Reader does not support OLE automation, except for the PDF browser controls provided in the AcroPDF object.

The best practical resources for Visual Basic or Visual C# programmers, besides the object browser, are the sample projects. The samples demonstrate use of the Acrobat OLE objects and contain comments describing the parameters for the more complicated methods. For more information see the ``Guide to SDK Samples``.

This chapter contains the following information:

| Topic | Description |
| -- | -- |
| OLE capabilities in Acrobat | Describes at a high level what you can do with OLE for interapplication communication. |
| Development environment considerations | Describes the benefits and drawbacks of using particular development environments and the required knowledge for each environment. |
| Using the Acrobat OLE interfaces | Explains the use of the CAcro and COLEDispatchDriver classes. |
| Using the JSObject interface | Explains the JSObject interface and provides examples of how it can be used. |
| Other development topics | Provides miscellaneous information about OLE automation. |
| Summary of OLE objects and methods | Provides a diagram of the OLE objects and methods and how they are related. |
| | |

For more information on OLE 2.0 and OLE automation, see the OLE Automation Programmer’s Reference, ISBN 1-55615-851-3, Microsoft Press. You can also find numerous articles at [http://msdn.microsoft.com](http://msdn.microsoft.com).

## OLE capabilities in Acrobat

For OLE automation, Acrobat provides three capabilities: rendering PDF documents, remotely controlling the application, and implementing PDF browser controls.

### On-screen rendering

You can render PDF documents on the screen in two ways:

- Use an interface similar to the Acrobat user interface.

In this approach, use the AVDoc object’s OpenInWindowEx method to open a PDF file in your application’s window. The window has vertical and horizontal scroll bars, and has buttons on the window’s perimeter for setting the zoom factor. Users interacting with this type of window find its operation similar to that of working in Acrobat. For example, links are active and the window can display any text annotation on a page.

The ActiveView sample in the Guide to SDK Samples shows how you can use this approach.

- Use the PDPage object’s DrawEx method.

In this approach, you provide a window and a device context, as well as a zoom factor. Acrobat renders the current page into your window. The application must manage the scroll bars and other items in the user interface.

The StaticView sample in the Guide to SDK Samples shows how you can use this approach.

### Remote control of Acrobat

You can control Acrobat remotely in two ways:

- Given the exported interfaces, you can write an application that manipulates various aspects of PDF documents, such as pages, annotations, and bookmarks. Your application might use AVDoc, PDDoc, PDPage, and annotation methods, and might not provide any visual feedback that requires rendering into its application window.

- You can launch Acrobat from your own application, which has set up the environment for the user. Your application can cause Acrobat to open a file, set the page location and zoom factor, and possibly even select some text. For example, this could be useful as part of a help system.

### PDF browser controls

You can use the AcroPDF library to display a PDF document in applications using simplified browser controls. In this case, the PDF document is treated as an ActiveX document, and the interface is available in Acrobat Reader.

Load the document with the AcroPDF object’s LoadFile method. You can then implement browser controls for the following functionality:

- To determine which page to display

- To choose the display, view, and zoom modes

- To display bookmarks, thumbs, scrollbars, and toolbars

- To print pages using various options

- To highlight a text selection

## Development environment considerations

You have a choice of environments in which to integrate with Acrobat: Visual Basic, Visual C#, and Visual C++.

If possible, use Visual Basic or Visual C#. The run-time type checking offered by the CreateObject call in Visual Basic allows quick prototyping of an application, and in both of these languages the implementation details are simplified.

For comparison, consider the following examples, in which you can see strings with "AcroExch.App" and strings with "Acrobat.CAcroApp". The first is the form for the external string used by OLE clients to create an object of that type. The second is the form that is included in developer type libraries.

This example shows a Visual Basic subroutine to view a given page of an open document:

- Viewing a page with Visual Basic

```vb
Private Sub myGoto(ByVal where As Integer)
   Dim app as Object, avdoc as Object, pageview as Object

   Set app = CreateObject("AcroExch.App")
   Set avdoc = app.GetActiveDoc
   Set pageview = avdoc.GetAVPageView
   pageview.Goto(where)
End Sub
```
The following example does the same, but in Visual C++:

- Viewing a page with Visual C++

```c++
void goto(int where)
{
   CAcroApp app;
   CAcroAVDoc *avdoc = new CAcroAVDoc;
   CAcroAVPageView pageview;
   COleException e;
   app.CreateDispatch("AcroExch.App");
   avdoc->AttachDispatch(app.GetActiveDoc, TRUE);
   pageview->AttachDispatch(avdoc->GetAVPageView, TRUE);
   pageview->Goto(where);
}
```

The next example shows how to use PDF browser controls to view a page in Visual Basic:

- Using AcroPDF browser controls with Visual Basic

```vb
Friend WithEvents AxAcroPDF1 As AxAcroPDFLib.AxAcroPDF
Me.AxAcroPDF1 = New AxAcroPDFLib.AxAcroPDF

'AxAcroPDF1

Me.AxAcroPDF1.Enabled = True
Me.AxAcroPDF1.Location = New System.Drawing.Point(24, 40)
Me.AxAcroPDF1.Name = "AxAcroPDF1"

Me.AxAcroPDF1.OcxState = CType(
      resources.GetObject("AxAcroPDF1.OcxState"),
      System.Windows.Forms.AxHost.State
)

Me.AxAcroPDF1.Size = New System.Drawing.Size(584, 600)
Me.AxAcroPDF1.TabIndex = 0
AxAcroPDF1.LoadFile("http://www.example.com/example.pdf")
AxAcroPDF1.setCurrentPage(TextBox2.Text)
```

The Visual Basic examples are simpler to read, write, and support, and the implementation details are similar to Visual C#.

In Visual C++, the CAcro classes hide much of the type checking that must be done. Using OLE automation objects in Visual C++ requires an understanding of the AttachDispatch and CreateDispatch methods of the COleDispatchDriver class. For more information, see ``Using the Acrobat OLE interfaces``.

>Note:
>
>The header files containing the values of constants that are required by C and C++ programmers to use OLE automation are located in the Acrobat DC SDK IAC directory. Visual Basic and Visual C# users do not need these header files, though it may be useful to refer to them in order to verify the constant definitions.

### Environment configuration

The only requirement for using the OLE objects made available by Acrobat is to have the product installed on your system and the appropriate type library file included in the project references for your project. The Acrobat type library file is named Acrobat.tlb. This file is included in the InterAppCommunicationSupport\Headers folder in the SDK. Once you have the type library file included in your project, you can use the object browser to browse the OLE objects.

It is not sufficient to install just an ActiveX control or DLL to enable OLE automation. You must have the full Acrobat product installed.

If you are a Visual Basic programmer, it is helpful to include the iac.bas module in your project (included in the headers folder). This module defines the constant variables.

### Necessary C knowledge

This guide and the ``Interapplication Communication API Reference`` describe the available objects and methods. These documents, as well as the API, were designed with C programming in mind and programming with the API requires some familiarity with C concepts.

Although you do not need the header files provided in the SDK, you can use them to find the values of various constants, such as **AV_DOC_VIEW**, that are referenced in the documentation. The file iac.h contains most of these values.

Some of the methods, such as **OpenInWindowEx**, can be initially confusing when used in Visual Basic. **OpenInWindowEx** takes a **long** for the **openflags** parameter. The options for this parameter, as provided in the ``Interapplication Communication API Reference``, are:

```
AV_EXTERNAL_VIEW — Open the document with the toolbar visible.
AV_DOC_VIEW — Draw the page pane and scrollbars.
AV_PAGE_VIEW — Draw only the page pane.
```

If you were developing in C, these strings would be replaced by a numeric value prior to compilation; passing these strings to the method would not raise an error. When programming in Visual Basic, these strings correspond to constant variables defined in iac.bas.

In some situations, you need to apply a bitwise OR to multiple values and pass the resultant value to a method. For example, in iac.h the **ntype** parameter of the **PDDocSave** method is a bitwise OR of the following flags:

```c
/* PDSaveFlags — used for PD-level Save 
** All undefined flags should be set to zero. 
** If either PDSaveCollectGarbage or PDSaveCopy are used, PDSaveFull must be used. */
typedef enum { 
   PDSaveIncremental = 0x0000,  /* write changes only */ 
   PDSaveFull = 0x0001,         /* write entire file */ 
   PDSaveCopy = 0x0002,         /* write copy w/o affecting current state */

   PDSaveLinearized = 0x0004,   /* write the file linearized for 
   **       page-served remote (net) access. */

   PDSaveBinaryOK = 0x0010, /* OK to store binary in file */

   PDSaveCollectGarbage = 0x0020  /* perform garbage collection on

   **       unreferenced objects */ 
} PDSaveFlags;
```

For example, if you would like to fully save the PDF file and optimize it for the Web (linearize it) within a Visual Basic application, pass **PDSaveFull** + **PDSaveLinearized** (both defined in iac.bas) into the **ntype** parameter; this is the equivalent of a binary **OR** of the **PDSaveFull** and **PDSaveLinearized** parameters.

In many instances, the numeric values are spelled out in comments in the Visual Basic sample code. However, knowledge of why the methods are structured in this way and how they are used in C can be useful to Visual Basic and Visual C# programmers.

## Using the Acrobat OLE interfaces

This section describes using the CAcro classes and the COleDispatchDriver class. The CAcro classes are subclasses of COleDispatchDriver.

### About the CAcro classes

OLE 2.0 support in Acrobat includes several classes whose names begin with “**CAcro**”, such as **CAcroApp** and **CAcroPDDoc**. Several files in the SDK encapsulate the definitions of these classes.

The **CAcro** classes are defined in the Acrobat type library acrobat.tlb. The **OLEView** tool in Visual Studio allows you to browse registered type libraries. Use acrobat.tlb when defining OLE automation for a project in Microsoft Visual C++. The files acrobat.h and acrobat.cpp are included in the Acrobat DC SDK, and implement a type-safe wrapper to the Acrobat automation server.

>Note:
>
>Do not modify the acrobat.tlb, acrobat.h, and acrobat.cpp files in the SDK; these define Acrobat’s OLE automation interface.

The **CAcro** classes inherit from the MFC **COleDispatchDriver** class. Understanding this class makes it easier to write applications that use the **CAcro** classes and their methods.

See the ``Interapplication Communication API Reference`` for details on the **CAcro** classes and their methods.

### About the COleDispatchDriver class

The **COleDispatchDriver** class implements the client side of OLE automation, providing most of the code needed to access automation objects. It provides the wrapper functions **AttachDispatch**, **DetachDispatch**, and **ReleaseDispatch**, as well as the convenience functions **InvokeHelper**, **SetProperty**, and **GetProperty**. You employ some of these methods when you use the Acrobat-provided automation objects. Other methods are used in the Acrobat implementation of these objects.

**COleDispatchDriver** is essentially a “class wrapper” for **IDispatch**, which is the OLE interface by which applications expose methods and properties so that other applications written in Visual Basic and Visual C# can use the application’s features. This provides OLE support for Acrobat applications.

### Using COleDispatchDriver objects and methods

This section discusses how to use the classes exported by acrobat.cpp, and shows when to call the **CreateDispatch** and **AttachDispatch** methods.

The following is a section of code from acrobat.h that declares the **CAcroHiliteList** class. **CAcroHiliteList** is a subclass of the **COleDispatchDriver** class, which means that it shares all the instance variables of **COleDispatchDriver**.

One of these variables is **m_lpDispatch**, which holds an **LPDISPATCH** for that object. An **LPDISPATCH** is a **long** pointer to an **IDispatch**, which can be considered an opaque data type representing a dispatch connection. **m_lpDispatch** can be used in functions that require an **LPDISPATCH** argument.

- CAcroHiliteList class declaration

```c++
class CAcroHiliteList : public COleDispatchDriver
{
public:
   CAcroHiliteList() {}        // Calls COleDispatchDriver default constructor
   CAcroHiliteList(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
   CAcroHiliteList(const CAcroHiliteList& dispatchSrc) :
      COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
   bool Add(short nOffset, short nLength);
};
```

The following is the related implementation section of the Add method from acrobat.cpp:

```c++
bool CAcroHiliteList::Add(short nOffset, short nLength)
{
   bool result;
   static BYTE parms[] =
      VTS_I2 VTS_I2;
   InvokeHelper(0x1, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
      nOffset, nLength);
   return result;
}
```

When the Add method is called, such as with this code from the following example ``Using the COleDispatchDriver class``,

```c++
hilite->Add(0, 10);
```

the **InvokeHelper** function is called. This **COleDispatchDriver** method takes a variable number of arguments. It eventually calls the Acrobat implementation for **CAcroHiliteList** object’s Add method. This happens across the virtual OLE “wires” and takes care of all the OLE details. The end result is that a page range is added to the **CAcroHiliteList** object.

The following is an implementation of a method adapted from the ActiveView sample:

- Using the COleDispatchDriver class

``` c++
// This code demonstrates how to highlight words with
// either a word or page highlight list

void CActiveViewDoc::OnToolsHilitewords()
{
   CAcroAVPageView pageView;
   CAcroPDPage page;
   CAcroPDTextSelect* textSelect = new CAcroPDTextSelect;
   CAcroHiliteList* hilite = new CAcroHiliteList;
   char buf[255];
   long selectionSize;

   if ((BOOL) GetCurrentPageNum() > PDBeforeFirstPage) {
      // Obtain the AVPageView
      pageView.AttachDispatch(m_pAcroAVDoc->GetAVPageView(),TRUE);

      // Create the Hilite list object
      hilite->CreateDispatch("AcroExch.HiliteList");
      if (hilite) {

   // Add the first 10 words or characters of that page to the highlight list
         hilite->Add(0,10);
         page.AttachDispatch(pageView.GetPage(), TRUE);

         // Create text selection for either page or word highlight list
         textSelect->AttachDispatch(page.CreateWordHilite(hilite->m_lpDispatch));
         m_pAcroAVDoc->SetTextSelection(textSelect->m_lpDispatch);
         m_pAcroAVDoc->ShowTextSelect();

         // Extract the number of words and the first word of text selection
         selectionSize = textSelect->GetNumText();
         if (selectionSize)
            sprintf (buf, "# of words in text selection: %ld\n1st word in text
               selection = '%s'", selectionSize, textSelect->GetText(0));
         else
            sprintf (buf, "Failed to create text selection.");

         AfxMessageBox(buf);
      }
   }

   delete textSelect;
   delete hilite;
}
```

In the preceding example, the objects with the prefix **CAcro** are all **CAcro** class objects—and they are also **COleDispatchDriver** objects—because all the Acrobat **CAcro** classes are subclasses of **COleDispatchDriver**.

*Instantiating a class is not sufficient to use it*. Before you use an object, you must attach your object to the appropriate Acrobat object by using one of the **Dispatch** methods of the **COleDispatchDriver** class. These functions also initialize the **m_lpDispatch** instance variable for the object.

This code from the previous example shows how to attach an **IDispatch** that already exists:

```c++
CAcroAVPageView pageView;
// Obtain the AVPageView 
pageView.AttachDispatch(m_pAcroAVDoc->GetAVPageView(), TRUE);
```

The **GetAVPageView** method of the **CAcroAVDoc** class returns an **LPDISPATCH**, which is what the **AttachDispatch** method is expecting for its first argument. The **BOOL** passed as the second argument indicates whether or not the **IDispatch** should be released when the object goes out of scope, and is typically **TRUE**. In general, when an **LPDISPATCH** is returned from a method such as **GetAVPageView**, you use **AttachDispatch** to attach it to an object.

The following code from the previous example uses the **CreateDispatch** method:

```c++
CAcroHiliteList *hilite = new CAcroHiliteList;
hilite->CreateDispatch("AcroExch.HiliteList");
hilite->Add(0, 10);
```

In this case, the **CreateDispatch** method both creates the **IDispatch** object and attaches it to the object. This code works fine; however, the following code would fail:

```c++
CAcroHiliteList *hilite = new CAcroHiliteList;
hilite->Add(0, 10);
```

This error is analogous to using an uninitialized variable. Until the **IDispatch** object is attached to the **COleDispatchDriver** object, it is not valid.

**CreateDispatch** takes a string parameter, such as "**AcroExch.HiliteList**", which represents a class. The following code is incorrect:

```c++
CAcroPDDoc doc = new CAcroPDDoc;
doc.CreateDispatch("AcroExch.Create");
```

This fails because Acrobat won’t respond to such a parameter. The parameter should be "**AcroExch.PDDoc**" instead.

The valid strings for CreateDispatch are as follows:

| Class | String |
| -- | -- |
| CAcroPoint | "AcroExch.Point" |
| CAcroRect | "AcroExch.Rect" |
| CAcroTime | "AcroExch.Time" |
| CAcroApp | "AcroExch.App" |
| CAcroPDDoc | "AcroExch.PDDoc" |
| CAcroAVDoc | "AcroExch.AVDoc" |
| CAcroHiliteList | "AcroExch.HiliteList" |
| CAcroPDBookmark | "AcroExch.PDBookmark" |
| CAcroMatrix | "AcroExch.Matrix" |
| AcroPDF | "AxAcroPDFLib.AxAcroPDF" |
|  |  |

Refer again to this code from the previous example:

```c++
CAcroPDPage page;
page.AttachDispatch(pageView.GetPage(), TRUE);
```

A **PDPage** object is required because the purpose of this code is to highlight words on the current page. Since it is a **CAcro** variable, it is necessary to attach to the OLE object before using its methods. **CreateDispatch** cannot be used to create a PDPage object because "**AcroExch.PDPage**" is not a valid string for **CreateDispatch**. However, the **AVPageView** method **GetPage** returns an **LPDISPATCH** pointer for a **PDPage** object. This is passed as the first argument to the **AttachDispatch** method of the page object. The **TRUE** argument indicates that the object is to be released automatically when it goes out of scope.

```c++
CAcroPDTextSelect* textSelect = new CAcroPDTextSelect;
textSelect->AttachDispatch
   (page.CreateWordHilite(hilite->m_lpDispatch));
m_pAcroAVDoc->SetTextSelection (textSelect->m_lpDispatch);
m_pAcroAVDoc->ShowTextSelect();
```

This code performs the following steps:
1. Declares a text selection object textSelect.

2. Calls the **CAcroPDPage** method **CreateWordHilite**, which returns an LPDISPATCH for a PDTextSelect. **CreateWordHilite** takes an LPDISPATCH argument representing a **CAcroHilite** list. The **hilite** variable already contains a **CAcroHiliteList** object, and its instance variable **m_lpDispatch** contains the LPDISPATCH pointer for the object.

3. Calls the **CAcroAVDoc** object’s SetTextSelection method to select the first ten words on the current page.

4. Calls the **AcroAVDoc**’s ShowTextSelect method to cause the visual update on the screen.

## Using the JSObject interface

Acrobat provides a rich set of JavaScript programming interfaces that can be used from within the Acrobat environment. It also provides the **JSObject** interface, which allows external clients to access the same functionality from environments such as Visual Basic.

In precise terms, **JSObject** is an interpretation layer between an OLE automation client, such as a Visual Basic application, and the JavaScript functionality provided by Acrobat. From a developer's point of view, programming **JSObject** in a Visual Basic environment is similar to programming in JavaScript using the Acrobat console.

This section explains how to extend Acrobat using JavaScript in a Visual Basic programming environment. It provides a set of examples to illustrate the key concepts.

Whenever possible, you should take advantage of these capabilities by using the **JSObject** interface available within the **AcroExch.PDDoc** object. To obtain the interface, invoke the object’s **GetJSObject** method.

### Adding a reference to the Acrobat type library

This procedure adds a reference to the Acrobat type library so that you can access the Acrobat automation APIs, including JSObject, in Visual Basic. Do this before using the JSObject interface, as in the examples that follow.

To add a reference to the Acrobat type library:

1. Install Acrobat and Visual Basic.

2. Create a new Visual Basic project from the Windows Application template. This provides a blank form and project workspace.

3. Select **Project** > **Add Reference** and click the **COM** tab.

4. From the list of available references, select **Adobe Acrobat 8.0 Type Library** and click **OK**.

### Creating a simple application

This example provides the minimum code to display “Hello, Acrobat!” in the Acrobat JavaScript console.

To set up and run the “Hello, Acrobat!” example:

1. Open the source code window for the default form by clicking View > Code.

2. Select (Form1 Events) from the selection box in the upper left corner of that window.

The selection box in the upper right corner now shows all the functions available to the Form1 object.

3. Select Load from the functions selection box. This creates an empty function stub. The Form1 Load function is called when Form1 is first displayed, so this is a good place to add the initialization code.

4. Add the following code to define some global variables before the subroutine.

```vb
Dim gApp As Acrobat.CAcroApp
Dim gPDDoc As Acrobat.CAcroPDDoc
Dim jso As Object
```
5. Add the following code to the private Form1_Load subroutine.

```vb
gApp = CreateObject("AcroExch.App")
gPDDoc = CreateObject("AcroExch.PDDoc")

If gPDDoc.Open("c:\example.pdf") Then
    jso = gPDDoc.GetJSObject
    jso.console.Show
    jso.console.Clear
    jso.console.println ("Hello, Acrobat!")
    gApp.Show
End If
```

6. Create a file called example.pdf at the root level of the C: drive.

7. Save and run the project.

When you run the application, Acrobat is launched, Form1 is displayed, and the JavaScript Debugger window is opened, displaying “Hello, Acrobat!”.

- Displaying “Hello, Acrobat!” in the JavaScript console

```vb
Dim gApp As Acrobat.CAcroApp
Dim gPDDoc As Acrobat.CAcroPDDoc
Dim jso As Object

Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs)
      Handles Me.Load
   gApp = CreateObject("AcroExch.App")
   gPDDoc = CreateObject("AcroExch.PDDoc")
   If gPDDoc.Open("c:\example.pdf") Then
      jso = gPDDoc.GetJSObject
      jso.console.Show
      jso.console.Clear
      jso.console.println ("Hello, Acrobat!")
      gApp.Show
   End If
End Sub
```

The Visual Basic program attaches to the Acrobat automation interface using the **CreateObject** call, and then shows the main window using the App object's **Show** command.

You may have a few questions after studying the code. For example, why is jso declared as an Object, while **gApp** and **gPDDoc** are declared as types found in the Acrobat type library? Is there a real type for **JSObject**?

The answer is no, **JSObject** does not appear in the type library, except in the context of the **CAcroPDDoc.GetJSObject** call. The COM interface used to export JavaScript functionality through JSObject is known as an IDispatch interface, which in Visual Basic is more commonly known simply as an “Object” type. This means that the methods available to the programmer are not particularly well-defined. For example, if you replace the call to

```vb
jso.console.clear
```

with

```vb
jso.ThisCantPossiblyCompileCanIt("Yes it can!")
```

the compiler compiles the code, but fails at run time. Visual Basic has no type information for **JSObject**, so Visual Basic does not know if a particular call is syntactically valid until run-time, and will compile any function call to a **JSObject**. For that reason, you must rely on the documentation to know what functionality is available through the **JSObject** interface. For details, see the ``JavaScript for Acrobat API Reference``.

You may also wonder why it is necessary to open a PDDoc before creating a **JSObject**. Running the program shows that no document appears onscreen, and suggests that using the JavaScript console should be possible without a **PDDoc**. However, **JSObject** is designed to work closely with a particular document, as most of the available features operate at the document level. There are some application-level features in JavaScript (and therefore in **JSObject**), but they are of secondary interest. In practice, a **JSObject** is always associated with a particular document.

When working with a large number of documents, you must structure your code so that a new **JSObject** is acquired for each document, rather than creating a single **JSObject** to work on every document.

### Working with annotations

This example uses the JSObject interface to open a PDF file, add a predefined annotation to it, and save the file back to disk.

To set up and run the annotations example:
1. Create a new Visual Basic project and add the Adobe Acrobat type library to the project.

2. From the Toolbox, drag the **OpenFileDialog** control to the form.

3. Drag a **Button** to your form.

![img](https://help.adobe.com/en_US/acrobat/acrobat_dc_sdk/2015/HTMLHelp/Acro12_MasterBook/IAC_DevApp_OLE_Support/addingbutton.gif)

4. Select **View** > **Code** and set up the following source code:

- Adding an annotation

```vb
Dim gApp As Acrobat.CAcroApp

Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As
      System.EventArgs) Handles MyBase.Load
   gApp = CreateObject("AcroExch.App")
End Sub

Private Sub Form1_Closed(Cancel As Integer)
   If Not gApp Is Nothing Then
      gApp.Exit
   End If
   gApp = Nothing
End Sub

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As
      System.EventArgs) Handles Button1.Click
   Dim pdDoc As Acrobat.CAcroPDDoc
   Dim page As Acrobat.CAcroPDPage
   Dim jso As Object
   Dim path As String
   Dim point(1) As Integer
   Dim popupRect(3) As Integer
   Dim pageRect As Object
   Dim annot As Object
   Dim props As Object
   
   OpenFileDialog1.ShowDialog()
   path = OpenFileDialog1.FileName
   
   pdDoc = CreateObject("AcroExch.PDDoc")
   If pdDoc.Open(path) Then
      jso = pdDoc.GetJSObject
      If Not jso Is Nothing Then
      
         ' Get size for page 0 and set up arrays
         page = pdDoc.AcquirePage(0)
         pageRect = page.GetSize
         point(0) = 0
         point(1) = pageRect.y
         popupRect(0) = 0
         popupRect(1) = pageRect.y - 100
         popupRect(2) = 200
         popupRect(3) = pageRect.y
         
         ' Create a new text annot
         annot = jso.AddAnnot
         props = annot.getProps
         props.Type = "Text"
         annot.setProps props
         
         ' Fill in a few fields
         props = annot.getProps
         props.page = 0
         props.point = point
         props.popupRect = popupRect
         props.author = "John Doe"
         props.noteIcon = "Comment"
         props.strokeColor = jso.Color.red
         props.Contents = "I added this comment from Visual Basic!"
         annot.setProps props
      End If
      pdDoc.Close
      MsgBox "Annotation added to " & path
   Else
      MsgBox "Failed to open " & path
   End If
   
   pdDoc = Nothing
End Sub
```

5. Save and run the application.

The code in the **Form_Load** and **Form_Closed** routines initializes and shuts down the Acrobat automation interface. More interesting work happens in the Command button's click routine. The first lines declare local variables and show the Windows Open dialog box, which allows the user to select a file to be annotated. The code then opens the PDF file's **PDDoc** object and obtains a **JSObject** interface to that document.

Some standard Acrobat automation methods are used to determine the size of the first page in the document. These numbers are critical to achieving the correct layout, because the PDF coordinate system is based in the lower-left corner of the page, but the annotation will be anchored at the upper left corner of the page.

The lines following the "**Create a new text annot**" comment do exactly that, but this block of code bears additional explanation.

First, **addAnnot** looks as if it is a method of **JSObject**, but the JavaScript reference shows that the method is associated with the **doc** object. You might expect the syntax to be **jso.doc.addAnnot**. However, **jso** is the **Doc** object, so **jso.addAnnot** is correct. All of the properties and methods in the **Doc** object are used in this manner.

Second, observe the use of **annot.getProps** and **annot.setProps**. The Annot object is implemented with a separate properties object, meaning that you cannot set the properties directly. For example, you cannot do the following:

```vb
annot = jso.AddAnnot
annot.Type = "Text"
annot.page = 0
...
```

Instead, you must obtain the properties object of **Annot** using **annot.getProps**, and use that object for read or write access. To save changes back to the original **Annot**, call **annot.setProps** with the modified properties object.

Third, note the use of **JSObject**'s color property. This object defines several simple colors such as red, green, and blue. In working with colors, you may need a greater range of colors than is available through this object. Also, there is a performance hit associated with every call to **JSObject**. To set colors more efficiently, you can use code such as the following, which sets the annot's **strokeColor** to red directly, bypassing the color object.

```vb
dim color(0 to 3) as Variant
color(0) = "RGB"
color(1) = 1#
color(2) = 0#
color(3) = 0#
annot.strokeColor = color
```

You can use this technique anywhere a color array is needed as a parameter to a **JSObject** routine. The example sets the colorspace to RGB and specifies floating point values ranging from 0 to 1 for red, green, and blue. Note the use of the # character following the color values. These are required, since they tell Visual Basic that the array element should be set to a floating point value, rather than an integer. It is also important to declare the array as containing Variants, because it contains both strings and floating point values. The other color spaces ("T", "G", "CMYK") have varying requirements for array length. For more information, refer to the **Color** object in the ``JavaScript for Acrobat API Reference``.

>Note:
>
>If you want users to be able to edit annotations, set the JavaScript property Collab.showAnnotsToolsWhenNoCollab to true.

### Spell-checking a document

Acrobat includes a plug-in that can scan a document for spelling errors. The plug-in also provides JavaScript methods that can be accessed using **JSObject**. In this example, you start with the source code from the example ``Adding an annotation`` and make the following changes:

- Add a List View control to the main form. Keep the default name ListView1 for the control.

- Replace the code in the existing Command1_Click routine with the following:

- Spell-checking a document

```vb
Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As
      System.EventArgs) Handles Button1.Click
   Dim pdDoc As Acrobat.CAcroPDDoc
   Dim jso As Object
   Dim path As String
   Dim count As Integer
   Dim i As Integer, j As Integer
   Dim word As Variant
   Dim result As Variant
   Dim foundErr As Boolean
   
   OpenFileDialog1.ShowDialog()
   path = OpenFileDialog1.FileName
   foundErr = False
   pdDoc = CreateObject("AcroExch.PDDoc")
   
   If pdDoc.Open(path) Then
      jso = pdDoc.GetJSObject
      If Not jso Is Nothing Then
         count = jso.getPageNumWords(0)
         For i = 0 To count - 1
            word = jso.getPageNthWord(0, i)
            If VarType(word) = vbString Then
               result = jso.spell.checkWord(word)
               If IsArray(result) Then
                  foundErr = True
                  ListView1.Items.Add (word & " is misspelled.")
                  ListView1.Items.Add ("Suggestions:")
                  For j = LBound(result) To UBound(result)
                     ListView1.Items.Add (result(j))
                  Next j
                  ListView1.Items.Add ("")
               End If
            End If
         Next i
         jso = Nothing
         pdDoc.Close
         
         If Not foundErr Then
            ListView1.Items.Add ("No spelling errors found in " & path)
         End If
      End If
   Else
      MsgBox "Failed to open " & path
   End If
   
   pdDoc = Nothing
End Sub
```

In this example, note the use of the Spell object’s **check** method. As described in the ``JavaScript for Acrobat API Reference``, this method takes a word as input, and returns a null object if the word is found in the dictionary, or an array of suggested words if the word is not found.

The safest approach when storing the return value of a **JSObject** method call is to use a Variant. You can use the **IsArray** function to determine if the Variant is an array, and write code to handle that situation accordingly. In this simple example, if the program finds an array of suggested words, it dumps them out to the List View control.

### Tips for translating JavaScript to JSObject

Covering every method available to JSObject is beyond the scope of this document. However, the ``JavaScript for Acrobat API Reference`` covers the subject in detail, and much can be inferred from the reference by keeping a few basic facts in mind:

- Most of the objects and methods in the reference are available in Visual Basic, but not all. In particular, any JavaScript object that requires the new operator for construction cannot be created in Visual Basic. This includes the Report object.

- The Annots object is unusual in that it requires JSObject to set and get its properties as a separate object using the getProps and setProps methods.

- If you are unsure what type to use to declare a variable, declare it as a Variant. This gives Visual Basic more flexibility for type conversion, and helps prevent runtime errors.

- JSObject cannot add new properties, methods, or objects to JavaScript. Due to this limitation, the **global.setPersistent** property is not meaningful.

- JSObject is case-insensitive. Visual Basic often capitalizes leading characters of an identifier and prevents you from changing its case. Don't be concerned about this, since JSObject ignores case when matching the identifier to its JavaScript equivalent.

- JSObject always returns values as Variants. This includes property gets as well as return values from method calls. An empty Variant is used when a null return value is expected. When JSObject returns an array, each element in the array is a Variant. To determine the actual data type of a Variant, use the utility functions **IsArray**, **IsNumeric**, **IsEmpty**, **IsObject**, and VarType from the Information module of the Visual Basic for Applications (VBA) library.

- JSObject can process most elemental Visual Basic types for setting properties and for and input parameters for method calls, including Variant, Array, Boolean, String, Date, Double, Long, Integer, and Byte. JSObject can accept Object parameters, but only when the Object is the result of a property get or method call to a JSObject. JSObject fails to accept values of type Error and Currency.

## Other development topics

This section contains a variety of topics related to developing OLE applications.

### Synchronous messaging

The Acrobat OLE automation implementation is based on a synchronous messaging scheme. When an application sends a request to Acrobat, the application processes that request and returns control to the application. Only then can the application send Acrobat another message. If your application sends one message followed immediately by another, the second message may not be properly received: instead of generating a server busy error, it fails with no error message.

For example, this can occur with the **AVDoc.OpenInWindowEx** method, where a large volume of information regarding drawing position and mouse clicks is exchanged, and with the usage of the **PDPage.DrawEx** method on especially complex pages. With the **DrawEx** method, the problem arises when a WM_PAINT message is generated. If the page is complex and the environment is multi-threaded, the application may not finish drawing the page before the application generates another WM_PAINT message. Because the application is single-threaded, multi-thread applications must handle this situation appropriately.

### MDI applications

Suppose you create a multiple document interface (MDI) application that creates a static window into which Acrobat is displayed using the **OpenInWindowEx** call, and this window is based on the **CFormView** OLE class. If another window is placed on top of that window and is subsequently removed, the Acrobat window does not repaint correctly.

To fix this, assign the Clip Children style to the dialog box template (on which **CFormView** is based). Otherwise, the dialog box erases the background of all child windows, including the one containing the PDF file, which wipes out the previously covered part of the PDF window.

### Event handling in child windows
When a PDF file is opened with **OpenInWindowEx**, Acrobat creates a child window on top of it. This allows the application to receive events for this window directly. However, an application must also handle the following events: **resize**, **key up**, and **key down**.

The following example from the ActiveView sample shows how to handle a resize event:

- Handling resize events

```c++
void CActiveViewVw::OnSize(UINT nType, int cx, int cy)
{
   CWnd* pWndChild = GetWindow(GW_CHILD);
   if (!pWndChild)
      return;
   CRect rect;
   GetClientRect(&rect);
   pWndChild->
      SetWindowPos(NULL,0,0,rect.Width,rect.Height,
               SWP_NOZORDER | SWP_NOMOVE);

   CView::OnSize(nType, cx, cy);
}
```

After sending the message to the child window, it also does a resize. This results in both windows being resized, which is the desired effect.

### Determining if an Acrobat application is running

Use the Windows FindWindow method with the Acrobat class name. You can use the Microsoft Spy++ utility to determine the class name for the version of the application.

### Exiting from an application

When a user exits from an application using OLE automation, Acrobat itself or a web browser displaying a PDF document can be affected:

- If no PDF documents are open in Acrobat, the application quits.

- If a web browser is displaying a PDF document, the display goes blank. The user can refresh the page to redisplay it.


## Using DDE

This chapter describes DDE support in Acrobat under Microsoft Windows. Although DDE is supported, you should use OLE automation instead of DDE whenever possible because DDE is not a COM technology.

For complete descriptions of the parameters associated with DDE messages, see the DDE sections of the ``Interapplication Communication API Reference``.

For all DDE messages, the service name is acroview, the transaction type is XTYPE_EXECUTE, and the topic name is control. The data is the command to be executed, enclosed within square brackets. The item argument in the DdeClientTransaction call is NULL.

The following example sets up a DDE message:

- Setting up a DDE message

```vb
DDE_SERVERNAME = "acroview";
DDE_TOPICNAME = "control";
DDE_ITEMNAME = "[AppHide()]";
```

The square bracket characters in DDE messages are mandatory. DDE messages are case-sensitive and must be used exactly as described.

To be able to use DDE messages on a document, you must first open the document using the DocOpen DDE message. You cannot use DDE messages to close a document that a user opened manually.

You can use NULL for pathnames, in which case the DDE message operates on the front document.

If more than one command is sent at once, the commands are executed sequentially, and the results appear to the user as a single action. You can use this feature, for example, to open a document to a certain page and zoom level.

Page numbers are zero-based: the first page in a document is page 0. Quotation marks are needed only if a parameter contains white space.

The document manipulation methods, such as those for deleting pages or scrolling, work only on documents that are already open.

## Using Apple Events

You can use several objects and events to develop Acrobat applications for Mac OS. Some of the objects and events in the Apple event registry are supported, as well as Acrobat-specific objects and events. Acrobat supports the following categories of Apple events:

| Category | Description |
| -- | -- |
| Required events | Events that the Finder sends to all applications. |
| Core events | Events that are common to a wide variety of applications, though not universally applicable to all applications. |
| Acrobat-specific events | Events that are specific to Acrobat. |
| Miscellaneous Apple events | Events that are not in one of the preceding categories. |
|  |  |

When programming for Mac OS, use AppleScript with Acrobat whenever possible. For Apple events that are not available through AppleScript, handle them with C or other programming languages. For a complete description of the parameters, see the ``Interapplication Communication API Reference``.

For information on Apple events supported by the Acrobat Search plug-in, see the ``Acrobat and PDF Library API Reference``. For information on other plug-ins supporting additional Apple events, see ``Overview``.

For more information on Apple events and scripting, see *Inside Macintosh: Interapplication Communication*, ISBN 0-201-62200-9, Addison-Wesley. The content of this document is currently available at [http://developer.apple.com/documentation/mac/IAC/IAC-2.html](http://developer.apple.com/documentation/mac/IAC/IAC-2.html).

For more information on the AppleScript language, see the AppleScript Language Guide, ISBN 0-201-40735-3, Addison-Wesley. The content of this document is currently available at [http://developer.apple.com/documentation/AppleScript/Conceptual/AppleScriptLangGuide/](http://developer.apple.com/documentation/AppleScript/Conceptual/AppleScriptLangGuide/).

For more information on the core and required Apple events, see the Apple event registry for Mac OS. This file is in the AppleScript 1.3.4 SDK, which is currently available at [http://developer.apple.com/sdk/](http://developer.apple.com/sdk/).