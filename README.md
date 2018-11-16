
# 简明Excel VBA

## 目录

- [x] 0x00 文档列表
- [x] 0x01 语法说明
- [x] 0x02 VBA界面介绍
- [x] 0x03 对象操作说明
- [ ] 0x04 Excel 相关常用操作
- [ ] 0x04 Trouble shooting
- [ ] 参考资料
- [x] Change log

## 0x00 文档列表
- [Excel-vba coding规约/开发规范](https://github.com/Youchien/development-specification/blob/master/doc/source/Excel-vba%20Language%20Specification.md)
- [Excel VBA 参考,官方文档,适用2013及以上](https://msdn.microsoft.com/zh-cn/library/ee861528.aspx)
- [Excel宏教程 (宏的介绍与基本使用)](http://blog.csdn.net/lyhdream/article/details/9060801)
- [Excel2010中的VBA入门,官方文档](https://docs.microsoft.com/zh-cn/previous-versions/office/ee814737(v=office.14))
- [Excel VBA的一些书籍资源,百度网盘](https://pan.baidu.com/s/1ktVmW63s8utBpAdcGnJfJA)  （提取码: `j92n`）
- [Excel 函数速查手册](https://support.office.com/zh-cn/article/Excel-%E5%87%BD%E6%95%B0%EF%BC%88%E6%8C%89%E7%B1%BB%E5%88%AB%E5%88%97%E5%87%BA%EF%BC%89-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb?ui=zh-CN&rs=zh-CN&ad=CN)
- [VBA的一些使用心得](http://www.cnblogs.com/techyc/p/3355054.html)
- [VBA函数参考](https://msdn.microsoft.com/zh-cn/library/office/jj692811.aspx)
- [VBA入门参考，英文](http://analystcave.com/vba-cheat-sheet/)

## 0x01 语法说明

都知道学会了英语语法，再加上大量的词汇基础，就算基本掌握了英语了。
类似的要使用vba，也要入乡随俗，了解他的构成，简单的说vba包含`数据类型`、
`变量`/`常量`、`对象`和常用的`语句结构`。

不过呢在量和复杂度上远低于英语，不用那么痛苦的记单词了，所以vba其实很简单的。
熟悉了规则之后剩下就是查官方函数啦，查Excel提供的可操作对象啦。

顺带一提的是，函数其实也很容易理解，方便使用。拿到一个函数，例如`Sum`，
只要知道它是求多个数的和就够了，剩下的就是用了。例如`Sum(1000,9)`结果就是`1009`了。
函数的一大好处就是隐藏具体实现细节，提供简洁的使用方法。


### 1.1 数据和数据类型

Excel里的每一个单元格都是一个`数据`，无论是数字、字母或标点都是数据。
对数据排排队，吃果果，对不同的数据扔到不同的篮子里归类，篮子就是`数据类型`了。

在Excel-vba中，`数据类型`只有`数值`、`文本`、`日期`、`逻辑`或`错误`五种类型。
前四种最为常用。具体描述参见下表：


| 类型 | 类型名称 | 范围 | 占用空间|声明符号 | 备注|
|--------|-------|-----|--------|-----|----|
| **逻辑型**|
| 布尔 | Boolean|逻辑值True或False|2|
|**数值型**|
|字节| Byte | 0~255的整数|1|
|整数| Integer| -32768~32767|2|%|
|长整数|Long|-2147483648~2147483647|4|&|
|单精度浮点|Single||4|!|
|双精度浮点|Double||4|#|
|货币|Currency||8|@|
|小数|Decimal||14|
|**日期型**|
|日期|Date|日期范围:100/1/1~9999/12/31|8|
|**文本型**|
|变长字符串|String|0~20亿||$|
|定长字符串|String|1~65400||
|**其他**|
|变体型|Variant(数值)|保存任意数值，也可以存储Error,Empty,Nothing,Null等特殊数值|
|对象|Object|引用对象|4|

表1.1 VBA数据类型

补充一点是，数组就像一筐水果，里面可以存不止一个数据。
他不是一个具体的数据类型，叫数据结构更合适些。

### 1.2 常量和变量

定义后不能被改变的量，就是`常量`；相反的`变量`就能修改具体值。

在vba里，使用一个`变量`/`常量`要先声明。

`常量`声明方法如下:
` Const 常量名称 As 数据类型 = 存储在常量中的数据`
例如:
```vba
Const PI As Single = 3.14 ' 定义一个浮点常量为PI，值为3.14
```

变量声明方法如下：
```vba
Dim 变量名 As 数据类型
```
变量名，必须字母或汉字开头，不能包含空格、句号、感叹号等。

数据类型，对应上面 ↑　表1.1里的那些

更多的声明方法，跟`Dim`声明的区别是作用范围不同：
```vba
Private v1 As Integer   ' v1为私有整形变量
Public v2 As String     ' v2为共有字符串变量
Static v3 As Integer    ' v3为静态变量，程序结束后值不变

' 变量声明之后，就可以赋值和使用了
v1 = 1009
v2 = "1009"
v3 = 1009

' 使用类型声明符，可以达到跟上面同样的效果
public v2$  ' 与 Public v2 As String 效果一样

' 声明变量时，不指定具体的类型就变成了Variant类型，根据需要转换数据类型
Dim v4
```

使用数组和对象时，也要声明，这里说下数组的声明：
```vba
' 确定范围的数组，可以存储b-a+1个数，a、b为整数
Dim 数组名称(a To b) As 数据类型

Dim arr(1 TO 100) As Integer ' 表示arr可以存储100个整数
arr(100) '表示arr中第100个数据

' 不指定a，直接声明时，默认a为0
Dim arr2(100) As Integer ' 表示arr可以存储101个整数,从0数
arr2(100) '表示arr2中第101个数据

' 多维数组
Dim arr3(1 To 3,1 To 3,1 To 3) As Integer ' 定义了一个三维数组，可以存储3*3*3=27个整数

' 动态数组，不确定数组大小时使用
Dim arr4() As Integer   ' 定义arr4为整形动态数组
ReDim arr4(1 To v1)     ' 设定arr4的大小，不能重新设定arr4的类型

```

除了用`Dim`做常规的数组的声明，还有下面这些声明数组的方式:
```vba
' 使用Array函数将已知的数据常量放到数组里
Dim arr As Variant        ' 定义arr为变体类型
arr = Array(1,1,2,3,5,8,13,21) ' 将整数存储到arr中,索引默认从0开始

' 使用Split函数分隔字符串创建数组
Dim arr2 As Variant
arr2 = Split("hello,world",",") ' 按,分隔字符串 hello,world 并赋值给arr2

' 使用Excel单元格区域创建数组
' 这种方式创建的数组，索引默认从1开始
Dim arr3 As Variant
arr3 = Range("A1:C3").Value   ' 将A1:C3中的数组存储到arr3中
Range("A4:C6").Value= arr3    ' 将arr3中的数据写入到A4:C6中的区域

```

//TODO 补充操作excel赋值的动图

**数组常用的函数**

|函数|函数说明|参数说明|示例|
|----|----|----|----|
|`UBound(Array arr,[Integer i])`|数组最大的索引值|`arr`:数组;`i`:整形,数组维数|
|`LBound(Array arr,[Integer i])`|数组最小的索引值|同上|
|`Join(Array arr,[String s])`|合并字符串|`arr`:数组;`s`:合并的分隔符|
|`Split(String str,[String s])`|分割字符串|`str`:待分割的字符串;`s`:分割字符串的分隔符|

> 函数说明
>
> UBound(Array arr,[Integer i]);
> UBound为函数名
> arr和i为UBound的的参数，用中括号括起来的表示i为非必填参数
> arr和i之前的Array,Integer表示对应参数的数据类型
>
> 补充
> [VBA 内置函数列表](https://msdn.microsoft.com/zh-cn/library/office/jj692811.aspx)

### 1.3 运算符

运算符的作用是对数据进行操作，像加减乘除等。这块不再具体说明，列一下vba中常用的运算符。

|运算符|作用|示例|
|----|----|----|
|**算术运算符**|
|+|求两个数的和|
|-|求两个数的差|
|*|求两个数的乘积|
|/|求两个数的商|
|`\`|求两个数相除后所得商的整数|
|^|求一个数的某次方|
|Mod|求两个数相除后所得的余数| 10 Mod 9=3|
|**比较运算符**|
|=|比较两个数据是否相等|相等返回 True;否则返回False|
|<>|不相等|
|<|小于|
|>|大于|
|<=|不大于|
|>=|不小于|
|Is|比较连个对象的引用关系|
|Like|比较两个字符串是否匹配| String1 Like String2|
|**文本运算符**|
|+|连接两个字符串|
|&|连接两个字符串|
|**逻辑运算符**|
|And|逻辑与|
|Or|逻辑或|
|Not|逻辑非|
|Xor|逻辑抑或|`表达式1 Xor 表达式2`两个表达式返回的值不相等时为True|
|Eqv|逻辑等价|`表达式1 Eqv 表达式2`两个表达式返回的值相等时为True|
|Imp|逻辑蕴含|

```vba
' Like是个比较有用的运算符，常用来做匹配或模糊匹配。
' 在模糊匹配的时候，有一些通配符能方便模糊匹配规则的书写
"这是一个demo1" Like "*demo1" = True '*号表示匹配任意多个字符
"这是一个demo2" Like "????demo2" = True '?号表示匹配任意单个字符
"这是一个demo3" Like "*demo#" = True '#号表示匹配任意数字
```

### 1.4 语句结构

程序通常都是顺序依次执行的。语句结构用来控制程序执行的步骤，一般有`选择`语句、`循环`语句。

**选择**
`选择`用来判断程序执行那一部分代码
```vba
'-----------------------------------
' If...Then...End If
' If选择可以嵌套使用
' 常用的三种形式
'-----------------------------------

' 普通模式
If 10>3 Then
    操作1'执行这一步
End If

' 增加Else
If 1>2 Then
      操作1
Else
    操作2'执行这一步
End If

'嵌套If
If 10 > 3 Then
    If 1 > 2 Then
        操作1
    Else
        操作2    ' 执行这一步
    End If
Else
    操作3
End If

' Select...Case... 多选一
Dim Length As Integer
Length=10
Select Length
    Case Is >=8
        操作1 '执行这一步
    Case Is >20
        操作2
    Case Else
        操作3
End Select
```

**循环**
`循环`用来让程序重复执行某段代码
```vba
' For...Next循环
' For 循环变量 = 初始值 To 终值 Step 步长
Dim i As Integer
For i = 1 To 10 Step 2 ' 设定i从1到10，每次增加2，总共执行5次
    操作1   ' 可以通过设定 Exit For 退出循环
Next i

' For Each..循环，又称遍历
' For Each 变量 In 集合或数组
Dim arr
Dim i As Integer
arr = Array(1, 2, 3, 4, 5)
For Each i In arr ' 定义变量i，遍历arr数组
    操作1
Next i

' Do...While循环
' Do While 表达式   表达式为假时跳出循环
Dim i As Integer
i = 1
Do While i < 5  ' 循环5次
    i = i + 1
Loop

'将判断条件后置的Do..While
Dim i As Integer
i = 1
Do
    i = i + 1
Loop While i<5 '循环4次

' Do Until 直到..循环
' Do Until 表达式    表达式为真时跳出循环
Dim i As Integer
i = 5
Do Util i < 1  
    i = i - 1
Loop

' 后置的Do Until
Dim i As Integer
i = 5
Do
    i = i - 1
Loop Util i<1  
```

`选择`和`循环`提供了多种实现同一目的的语句结构，他们都能实现同样的作用，
差别一般是初始条件。还有书写的复杂度。正确的选择要使用的语句结构，
代码逻辑上会更清楚，方便人的阅读。

**简写**

在操作对象的属性时常常要先把对象调用路径都写出来，用`with`可以简化这一操作
```vba
' 简化前
WorkSheets("表1").Range("A1").Font.Name="仿宋"
WorkSheets("表1").Range("A1").Font.Size=12
WorkSheets("表1").Range("A1").Font.ColorIndex=3

' 使用`With`
With WorkSheets("表1").Range("A1").Font
    .Name = "仿宋"
    .Size = 12
    .ColorIndex =3
End With
```

### 1.5 过程和函数

`Sub`和`Function`是VBA提供的两种封装体，利用宏录制器得到的就是`Sub`。
两者的区别不大，`Sub`不需要返回值，`Function`可以定义返回值和返回的类型。

**Sub**
```vba
[Private|Public] [Static] Sub 过程名([参数列表 [As 数据类型]])
    [语句块]
End Sub
' [Private|Public]定义过程的作用范围
' [Static]定义过程是否为静态
' [参数列表]定义需要传入的参数
```
调用`Sub`的方法有三种，使用`Call`、直接调用和`Application.Run`

举个栗子：
![Alt text](/doc/source/images/1505555701907.png)

**Function**

vba内部提供了大量的函数，也可以通过`Function`来定义函数，实现个性化的需求。
```vba
[Public|private] [Static] Function 函数名([参数列表 [As 数据类型]]) [As 数据类型]
    [语句块]
    [函数名=过程结果]
End Function
```
使用函数完成上面的栗子：
![Alt text](/doc/source/images/1505556598033.png)

**参数传递**

参数传递的方式有两种，引用和传值。
传值，只是将数据的内容给到函数，不会对数据本身进行修改。
引用，将数据本身传给函数，在函数内部对数据的修改将同样的影响到数据本身的内容。

参数定义时，使用`ByVal`关键字定义传值，子过程中对参数的修改不会影响到原有变量的内容。
默认情况下，过程是按引用方式传递参数的。在这个过程中对参数的修改会影响到原有的变量。
也可以使用`ByRef`关键字显示的声明按引用传参。
```vba
Sub St1(ByVal n As Integer,ByRef range)
	...Other code
End SUb
```


### 1.6 补充

- 在vba中使用 `'`进行代码注释
- 在很长的语句中使用`_`来分割成多行
- 在有很多嵌套判断中，代码的可读性会变得很差，一般讲需要返回的内容及时返回，减少嵌套
- `Sub`中默认按引用传递参数，所以注意使用，一般不要对外面的变量进行修改，将封装保留在内部


- `Dim`和`Set`的区别  [参考](http://blog.csdn.net/nctu_to_prove_safety/article/details/53148962)

```vba
' vba中使用Dim设定变量类型，Set将对象引用赋值给变量

' 栗子
' 将Range对象赋值给变量rg
Dim rg As Range  ' 声明rg为Range对象
Set rg = Range("A1") ' 设定rg为Range("A1")的引用，之后操作rg和操作Range("A1")一样了

' 如果不使用Set，下面的代码将报错
Dim rg As Range
rg = Range("A1")   ' 这段代码将报错

' 在非显示声明rg的前提下，下面的代码将会得到不一样的结果
rg = Range("A1")  ' rg将会是Range("A1")的内容，rg的类型将会是一种基本类型，Integer/String等
Set rg = Range("A1")   ' 这种情况下，rg将会是Range对象
```


### 1.7示例

举个排序的栗子，要对`A1:A20`的单元格区域进行排序，区域内的内容为1-100的随机整数，
规则是大于50的倒序排列，小于50的正序排列。将结果显示在`B1:B20`的区域里。

在这个栗子中，首先定义一个`Sub`过程来随机生成`A1:A20`区域的内容。
代码如下:

![Alt text](/doc/source/images/demo1.1.gif)

```vba
' 创建随机整数，并赋值
Sub createRandom(times As Integer)
    Dim num As Integer
    Dim arr() As Integer
    ReDim arr(times)

    For num = 1 To times
        Randomize (1) ' 初始化随机数
        arr(num) = Rnd(1) * 10000 \ 100 ' Rnd随机数函数生成0~1的浮点数
        ' 上面使用了运算符进行取整，也可以根据需求使用vba内部的取整函数达到同样的效果
        ' arr(num) = Int(Rnd(1) * 100)
        ' arr(num) = Round(Rnd(1) * 100)
        Range("A" & num) = arr(num)
    Next num
End Sub

'自定义排序
Function defSort(rgs) As Variant
  Dim arr() As Integer
  Dim total As Integer
  Dim rg
  Dim st As Integer '数组开始标记
  Dim ed As Integer '数组结束标记

  Debug.Print "rgs类型:"; TypeName(rgs)
  total = UBound(rgs)
  ReDim arr(total)
  st = 1
  ed = total

  '对数组分区
  For Each rg In rgs
    If rg > 50 Then
      arr(ed) = rg
      ed = ed - 1
    Else
      arr(st) = rg
      st = st + 1
    End If
  Next rg

  Dim i As Integer
  Dim j As Integer
  Dim tmp As Integer

  '冒泡排序
  For i = 1 To total
    For j = i To total
      If arr(i) > 50 And arr(j) > 50 Then '大于50的倒序排列
        If arr(i) < arr(j) Then
          tmp = arr(i)
          arr(i) = arr(j)
          arr(j) = tmp

          Debug.Print "大于50的"; i; j; tmp '程序运行过程中在立即窗口显示执行内容，用于调试程序
        End If
      ElseIf arr(i) <= 50 And arr(j) <= 50 Then '小于50的正序排列
        If arr(i) > arr(j) Then
          tmp = arr(i)
          arr(i) = arr(j)
          arr(j) = tmp

          Debug.Print "不大于50的"; i; j; tmp
        End If
      Else
        Exit For
      End If
    Next j
  Next i
  defSort = arr
End Function


'程序入口
Sub main()
  Const SORT_NUM = 20
  Dim rgs
  Dim arr

  createRandom SORT_NUM '初始化待排序区域

  rgs = range("A1:A" & SORT_NUM)
  arr = defSort(rgs)

  '循环赋值
  For i = 1 To SORT_NUM
    range("B" & i) = arr(i)
  Next i
End Sub
```

## 0x02 VBA界面介绍
### 2.1 整体界面说明
![Alt text](/doc/source/images/1505749555407.png)

### 2.2 工程资源管理器（Project Explore）说明

显示快捷键：`Ctrl + R`，也可以点击菜单栏 View -> <u>P</u>roject Explore 显示。
在一个VBA项目中，实际可以在5个代码模块中书写VBA代码，如下图所示：

![Alt text](/doc/source/images/vba_code_modules.png)

1. Code Modules – Code Modules是我们存储宏的最常见的地方。
模块位于工作簿中的 `Modules` 文件夹中。

2. Sheet Modules – 工作簿中的每个工作表在Microsoft Excel Objects文件夹中
都有一个工作表对象。双击sheet对象就会打开它的代码模块，我们可以在其中添加事件过程(宏)。
这些宏在用户执行表单中的特定操作时运行。比如如下code：
如果在该sheet中的选择位置发生改变，就会*自动执行* `Worksheet_SelectionChange` 方法，
选择所选单元格的整个行和列。

```VBA
Private Sub Worksheet_SelectionChange(ByVal Target As Range) 'Worksheet_SelectionChange
    Application.EnableEvents = False

    With Target
        Union(.EntireRow, .EntireColumn).Select
        .Activate
    End With

    Application.EnableEvents = True
End Sub
```

3. ThisWorkbook Module – 每个工作簿都包含一个 `ThisWorkbook` 对象，
其总是位于和工作表对象相同的文件夹(Microsoft Excel Objects)内的最底部。
我们可以在这个工作簿中运行基于事件的宏。

4. Userforms – 做过VB项目的人对这个应该不会陌生。在这个模块下我们可以创建Windows窗体，
进行图形化交互。在这个模块写的code大部分都是和win窗体相关的代码。

5. Class Modules – 在`Class Modules`文件夹中，允许我们编写宏来创建对象、属性和方法。
当我们想要创建对象库中不存在的自定义对象或集合时，可以使用该类模块。

**总结**：`Modules`、 `ThisWorkbook`、 `Sheet` 三者区别：

`Modules` 是相似功能和子程序的集合，通常根据功能进行分组。

`ThisWorkbook` 是Workbook对象的私有模块。
例如， Workbook_Open()， Workbook_Close() 例程驻留在此模块中。
（[工作簿对象参考](https://docs.microsoft.com/zh-cn/office/vba/api/excel.workbook)）

`Sheet1`，`Sheet2` 是单个工作表的私有模块。在它们中，您将会放入该表的特定功能。
例如：`Worksheet_Activate` ， `Worksheet_Deactivate` ， `Workbook_SheetChange`
是提供给的默认事件，这样你就可以在各自的私有工作表模块中处理它们。
（[工作表对象参考](https://msdn.microsoft.com/en-us/library/office/ff847327.aspx)）

在模块里使用Cells、range等时表示的是当前激活的工作表；而在sheet里面写的话，
为当前工作表里的cells，如果你在sheet1代码里要引用其他工作表的话，不能这样。

```vba
sheet2.select
cells(1,1) = 1
```

因为你的代码在sheet1下，cells就一定是sheet1的
另外，在sheet下面可以使用Me，表示自身
如sheet1.visible = False，可以简化为: Me.visible = False

如果一个Funtion是在`Modules`里定义的，那么就可以在任意的Worksheet里调用，
但如果只是在Worksheet里定义的Funtion，其他的Worksheet是调用不了的。
也就是说，模块（Modules）是公共的地方。

### 2.3 设置VBA Macro Project 密码保护

![Alt text](/doc/source/images/password_protect_setting.png)

在VBA界面依次点击：<u>T</u>ools -> VBAProject Prop<u>e</u>rties ->
Projection 界面设置


### 2.4 常用快捷栏及窗口设置
默认情况下某些常用的窗口VBA界面是不显示的，比如立即窗口，编辑操作捷栏（批量注释取消等）

#### 2.4.1 显示编辑栏
鼠标右键点击空白的快捷栏位置，勾选 `Edit` 选项会显示出如下快捷栏

![Alt text](/doc/source/images/toolbars_edit_setting.png)

#### 2.4.2 显示立即窗口(Immediate window)
显示快捷键： `Ctrl + G`，也可以点击菜单栏 View -> <u>I</u>mmediate window 显示。


## 0x03 对象操作说明
Excel中的每个单元格，工作簿都是可以操作的对象；可以对对象进行复制、粘贴、删除等，
也可操作对象的各种属性，来控制其展示和行为。

在Excel中，对象有不同的层级关系:

![Alt text](/doc/source/images/1505548045994.png)

实际上Excel中可操作的对象远不止这些，具体的可以参考
[Excel 对象模型](https://msdn.microsoft.com/zh-cn/library/office/ff194068.aspx)

类似于数组，将各种类型的对象封装到一块可以组成集合。
一个集合中调用对象的例子：
![Alt text](/doc/source/images/1505548422147.png)



### 3.1 对象简述

对象一般包含下面三种特性：

- 属性

属性表示对象的特征，一般为名词。例如`Workbook.ActiveSheet`表示工作簿当前
处于激活状态的工作表对象。

- 方法

方法表示对象可用的操作或可执行的动作。例如`Workbook.Activate`表示
激活工作簿的第一个工作表。

- 事件

事件表示对象可以被触发的行为，一般触发后会执行对应的代码。
例如`Workbook.Activate`表示工作簿中的工作表被激活了，然后执行对应的方法。

下面的代码就是在`Workbook`被打开时，将工作簿最大化的例子。

```vba
Private Sub Workbook_Open()
    Application.WindowState = xlMaximized
End Sub
```


VBA中有很多对象，常用的对象如下:

|对象|对象说明| 文档地址|
|----|----|----|
|Application|代表Excel应用程序|[文档](https://msdn.microsoft.com/zh-cn/library/ff194565.aspx)|
|Workbook|代表Excel的工作簿|[文档](https://msdn.microsoft.com/zh-cn/library/ff835568.aspx)|
|Worksheet|代表Excel的工作表|[文档](https://msdn.microsoft.com/zh-cn/library/ff194464.aspx)|
|Range|代表Excel的单元格，可以是单个单元格或单元格区域|[文档](https://msdn.microsoft.com/zh-cn/library/office/ff838238.aspx)|


### 3.2 Application对象



### 3.3 Range对象
![Alt text](/doc/source/images/1505548886377.png)

![Alt text](/doc/source/images/1505549069568.png)


## 0x04 Excel 相关常用操作

### 4.1 打开Excel两种方式

- 利用 `GetObject` 方法打开Excel文档
```vba
    Sub GetWorkbook()
        Dim wbWorkFile As Workbook
        Set wbWorkFile = GetObject("D:\test.xlsx")
        ' wbWorkFile.Windows(1).Visible = True ' 这种方法打开的文件是隐藏的，如果需要显示，则设置Visible值为ture
        wbWorkFile.Close False
        Set wbWorkFile = Nothing
    End Sub
```

- 利用 `Open` 方法打开Excel文档
```vba
Sub OpenWorkbook()
    Dim wbWorkFile As Workbook
    Set wbWorkFile = Workbooks.Open("D:\test.xlsx")
    wbWorkFile.Windows(1).Visible = False
    wbWorkFile.Close False
    Set wbWorkFile = Nothing
End Sub
```

延伸其扩展方法：
- GetObject封装方法，可以作为共通function
```vba
Sub GetWorkbook()
    Dim objExcel                As Object       ' 用于存放Microsoft Excel 引用的变量。
    Dim blnExcelWasNotRunning   As Boolean      ' 用于最后释放的标记。

    ' 测试 Microsoft Excel 的副本是否在运行。
    On Error Resume Next                        ' 延迟错误捕获。
    ' 不带第一个参数调用 Getobject 函数将返回对该应用程序的实例的引用。如果该应用程序不在运行，则会产生错误。
    Set objExcel = Getobject(, "Excel.Application")
    If Err.Number <> 0 Then blnExcelWasNotRunning = True
    Err.Clear                                   ' 如果发生错误则要清除 Err 对象。

    Set objExcel = Getobject("C:\excel.xlsx")   ' 将对象变量设为对要看的文件的引用。

    ' 设置其 Application 属性，显示 Microsoft Excel。然后使用 objExcel 对象引用的 Windows 集合显示包含该文件的实际窗口。
    objExcel.Application.Visible = True
    objExcel.Parent.Windows(1).Visible = True
    ' 在此处对文件进行操作。
    ' ...
    ' 如果在启动时，Microsoft Excel 的这份副本不在运行中，则使用 Application 属性的 Quit 方法来关闭它。
    ' 注意，当试图退出 Microsoft Excel 时，标题栏会闪烁，并显示一条消息询问是否保存所加载的文件。
    If blnExcelWasNotRunning = True Then
        objExcel.Application.Quit
    End IF

    Set objExcel = Nothing   ' 释放对该应用程序

End Sub
```

- OpenWorkbook封装方法，可以作为共通function
```vba
Function OpenWorkbook(ByVal strWorkbookFilePath As String)
    Dim wb As Workbook
    Dim fileName As String
    fileName = Dir(strWorkbookFilePath)

    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(strWorkbookFilePath)
    End If

    Set OpenWorkbook = wb

End Function
```


## Change log

- 2017/09/22  Fix对象操作说明的一些表述；补充追加 界面介绍及Excel相关常用操作
- 2017/09/22  Fix对象的一些表述错误
- 2017/09/18  补充`Set`&`Dim`;VBA界面介绍
- 2017/09/16  VBA语法说明
