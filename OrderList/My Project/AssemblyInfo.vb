Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security
Imports Microsoft.Office.Tools.Excel

' 有关程序集的常规信息通过下列属性集
' 控制。更改这些属性值可修改
' 与程序集关联的信息。

' 查看程序集属性的值

<Assembly: AssemblyTitle("OrderList")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("Microsoft")> 
<Assembly: AssemblyProduct("OrderList")> 
<Assembly: AssemblyCopyright("Copyright © Microsoft 2014")> 
<Assembly: AssemblyTrademark("")> 

' 将 ComVisible 设置为 false 使此程序集中的类型
' 对 COM 组件不可见。如果需要从 COM 访问此程序集中的类型，
' 则将该类型上的 ComVisible 属性设置为 true。
<Assembly: ComVisible(False)>

'如果此项目向 COM 公开，则下列 GUID 用于类型库的 ID
<Assembly: Guid("a08e590a-7a0c-48a1-b126-222ce2bb5516")> 

' 程序集的版本信息由下面四个值组成:
'
'      主版本
'      次版本 
'      内部版本号
'      修订号
'
' 可以指定所有这些值，也可以使用“内部版本号”和“修订号”的默认值，
' 方法是按如下所示使用“*”:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

' ExcelLocale1033 属性用于控制传递给 Excel 对象模型的区域设置。
' 将 ExcelLocale1033 设置为 true 会使 Excel 对象模型在
' 所有区域设置下采取同一行为，该行为与 Visual Basic for  Applications
' 的行为相符。将 ExcelLocale1033 设置为 false 会使 Excel 对象模型
' 的行为随用户的区域设置不同而不同，该行为
' 与 Visual Studio Tools for Office 版本 2003 的行为相符。这会导致
' 区分区域设置的信息(如公式名和日期格式)中出现意外结果。

<Assembly: ExcelLocale1033(True)>

<Assembly: SecurityTransparent()>
