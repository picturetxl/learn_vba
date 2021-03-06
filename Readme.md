# EXCEL - VBA 

> visual basic application --> VBA

目录

[TOC]



## 写在开头

### 课时安排

> 每周二上课

| check              |     时间     |               内容               |      地点       | 课时 |  实际人数   |
| ------------------ | :----------: | :------------------------------: | :-------------: | :--: | :---------: |
| :heavy_check_mark: | 14::00-14:45 |    VBA开发环境搭建和入门案例     | ERP研究所会议室 |  1   | 郑,张,温,衣 |
| :heavy_check_mark: | 14::00-15:30 | VBA操作工作表,工作簿,单元格对象1 | ERP研究所会议室 |  2   | 郑,张,温,衣 |
| :heavy_check_mark: | 14::00-15:30 | VBA操作工作表,工作簿,单元格对象2 | ERP研究所会议室 |  2   |  郑,张,温   |
|                    | 14::00-15:30 |        VBA事件和典型应用         | ERP研究所会议室 |  2   |             |
|                    | 14::00-14:45 |       VBA中使用函数和公式        | ERP研究所会议室 |  1   |             |
|                    | 14::00-15:30 |   VBA自定义函数和传参实现复用    | ERP研究所会议室 |  2   |             |
|                    | 14::00-14:45 |       VBA控件和窗体的使用        | ERP研究所会议室 |  1   |             |
|                    | 14::00-14:45 |         VBA用户信息交互          | ERP研究所会议室 |  1   |             |

### 教学目标

>  熟练掌握VBA,直接理解BPC内部的VBA代码并学会调试代码.使用VBA技巧给日常工作带来方便,提高工作效率. 

### 网上的文档

[w3school]( https://www.w3cschool.cn/excelvba/ )

### 记笔记的软件

[typora]( https://www.typora.io/ )



### 文档约定

|          符号           |       意义       | 备注 |
| :---------------------: | :--------------: | :--: |
|         :happy:         | 表示vb的语法相关 |      |
|        :warning:        | 过程中出现的警告 |      |
|      :red_circle:       | 过程中出现的错误 |      |
|        :hammer:         |      小技巧      |      |
| :ballot_box_with_check: |     解决方案     |      |
|                         |                  |      |
|                         |                  |      |
|                         |                  |      |
|                         |                  |      |





## 环境搭建

#### 打开开发工具

##### 1. 文件

![image-20200303091227778](Readme.assets/image-20200303091227778.png)

##### 2. 选项

![image-20200303091259040](Readme.assets/image-20200303091259040.png)

##### 3. 自定义功能区

![image-20200303091331384](Readme.assets/image-20200303091331384.png)



##### 4. 完成



![image-20200303091415450](Readme.assets/image-20200303091415450.png)



##### 5. 完整演示

![vba01](Readme.assets/vba01.gif)



### 启用宏

#### 1. 文件

#### 2. 选项

#### 3.信任中心 -> 信任中心设置

![image-20200303093400922](Readme.assets/image-20200303093400922.png)

#### 4. 宏设置->启用所有宏

![image-20200303093337165](Readme.assets/image-20200303093337165.png)



#### 5. 完整演示

![vba02](Readme.assets/vba02.gif)











## Start

+ 工资条案例

![image-20200303091047573](Readme.assets/image-20200303091047573.png)

### 录制宏

#### 辅助工具--录制宏

> 用于不常用功能或者复杂功能的代码书写
>
> 因为不常用的功能代码不知道怎么写,可以百度,也可以利用VBA的宏录制功能替代.

#### demo

> 所谓录制宏,相当于SAP的录屏功能.操作的每一步都会被记录下来.--所以不要乱点,想清楚步骤.
>
> 演示下获取把单元格变色的功能代码

+ 手工的操作步骤
  + 选中要变色的单元格
  + 变色



+ 录制宏

  1. 点击录制宏

     ![image-20200303091752967](Readme.assets/image-20200303091752967.png)

  2. 对宏取名->确定

     > 取名为change_color

     ![image-20200303092721591](Readme.assets/image-20200303092721591.png)

  3. 选中单元格

  4. 改颜色

     ![image-20200303093600234](Readme.assets/image-20200303093600234.png)

  5. 停止录制

  6. 查看生成的代码

     > 单击visual basic

     ![image-20200303094857228](Readme.assets/image-20200303094857228.png)

     > 模块->模块1

     ![image-20200303094749187](Readme.assets/image-20200303094749187.png)

     > 代码

     ```vb
     Sub change_color()
     '
     ' change_color 宏
     '
     
     '
         Range("C4").Select
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 65535
             .TintAndShade = 0
             .PatternTintAndShade = 0
         End With
     End Sub
     
     ```

  7. 完整实例

     <img src="Readme.assets/vba03.gif" alt="vba03"  />





## 数据类型

#### 整型integer

```vba
dim i as Integer
```



#### 单精度浮点数single

```vba
dim i as single
```



#### 双精度浮点数 double

```vba
dim i as double
```



#### 工作表类型

```vba
Dim sht As Worksheet '对象也是一种类型
```





## :bug: Bug

### 溢出

> 可能原因是整数溢出,可以更换类型为浮点数double



## 对象



### 工作表对象

#### 添加工作表

```vba
Sheets.Add after:=Sheets(Sheets.Count) '在表最后添加表
Sheets(Sheets.Count).Name = Sheet1.Range("a1").Value '改表名
```

#### 访问某一个表



+ 注意:

  > index 是从1开始 看到的excel工作表就是从第一个开始,依次往后,不管工作表的大名还是小名--> **sheets(index )**
  >
  > 而vba编辑器中的sheet4 是工作表的大名 2叫小名

  ![image-20200310150702189](Readme.assets/image-20200310150702189.png)



+ 选择工作表



1. 索引访问

   ```vba
   Sheets(index)
   ```

   

2. 直接访问

   ```vba
   sheet1
   ```

   ![image-20200323161647409](Readme.assets/image-20200323161647409.png)

#### 选中工作表

```vba
Sheet1.Select
```

#### 增加工作表

> 添加 在哪里添加 并给他一个表名

```vba
Sheets.Add after:=Sheets(Sheets.Count)
Sheets(Sheets.Count).Name = Sheet1.Cells(i, l)'cell(行,列)
```



### 工作簿对象

> 另存为saveas
>
> 关闭close

```vba
For Each sht In Sheets
    sht.Copy
    ActiveWorkbook.SaveAs Filename:="d:\data\" & sht.Name & ".xlsx"
    ActiveWorkbook.Close
Next
```









## 常见功能代码



### 自动筛选

> 第一步:选中需要筛选的区域,第二步设置过滤的条件 
>
> Field:说明按照第几列筛选
>
> Criteria1:说明该列的第几个值筛选

```vba
Sheet1.Range("a1:f1048").AutoFilter Field:=4, Criteria1:=Sheets(i).Name
```

> 恢复自动筛选状态

```vba
Sheet1.Range("a1:f1048").AutoFilter
```



### 输入输出对话框

#### 输入对话框

```vba
l = InputBox("请输入你要按哪列分") '保存在l变量里面
```

#### 输出对话框

```vba
MsgBox "已处理完毕"
```



### 拷贝数据

> 拷贝到哪里

```vba
Sheet1.Range("a1:f" & irow).Copy Sheets(j).Range("a1")
```



### 提示关闭

> 成对出现

```vba
Application.DisplayAlerts = False '成对出现的

	'....code

Application.DisplayAlerts = True
```



### 获取行数和列数

```vba
icolumn = Sheet1.Range("IV1").End(xlToLeft).Column '获取列数
irow = Sheet1.Range("a65536").End(xlUp).Row '获取行数
```

### 删除选区内容

```vba
Sheet1.Range("a1:f65536").ClearContents
```

### 删除整行

```vba
Range("a10").EntireRow.Delete
```





## 语法

###  :happy: 宏代码的结构

```vb
Sub 宏名()

    
End Sub
```

###  :happy: 定义变量

```vb
dim counter as Integer 
```



###  :happy: 循环

#### for循环

```vb
dim counter as Integer 
For counter = 1 To 5 step 2
'循环体
next

```

#### for each循环

```vba
For Each sht In Sheets
    If sht.Name = Sheet1.Range("a1") Then
    	k = 1
    End If
Next
```



### :happy: 判断

```vb
If 条件1 Then
	条件1为真时要执行的语句
ElseIf 条件2 Then
	条件2为真时要执行的语句
ElseIf 条件3 Then
	条件3为真时要执行的语句
ElseIf 条件N Then
	条件N为真时要执行的语句
Else
	所有条件都为假时要执行的语句
End If
```



### :happy:常见对象/属性

> 所谓对象就是干活的人.你有一件事你不会做,你要指定一个会做的人去做,这个人就是对象.
>
> 属性就是这个对象能给你提供什么.



## 快捷键

+ `Tab` : 向后缩进
+ `shift` + `tab`:向前缩进
+ `ctrl`+ `s` : 保存
+ `F8`:单步调试

### 制作工资条

> 要将第一行的entry复制到第一行一下的每一行

![image-20200303101109701](Readme.assets/image-20200303101109701.png)

​	

:question:如何实现

> 首先手工可以完成,但是效率太低.如果员工人数过多,那浪费太多的时间

+ 分解步骤
  + 选中第一行,复制
  + 插入复制的行
  + 循环做10次

+ 使用录制宏->找到每个子功能

  + 选中第一行,复制

    ```vb
     Rows("1:1").Select
     Selection.Copy
    ```

    ![image-20200303101617776](Readme.assets/image-20200303101617776.png)

  + 插入复制的行--> 使用**相对引用**

    ```vb
    ActiveCell.Offset(-6, 0).Rows("1:1").EntireRow.Select
    Selection.Copy
    ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown
    ```

  + :happy: 循环

    ```vb
    dim counter as Integer 
    For counter = 1 To 5 step 1
    '循环体
    next
    
    ```

  + :happy: 完整代码

    ```vb
    
    Sub gzt()
        ' 生成工资条
        Dim i As Integer ' 定义变量
    
        Rows("1:1").Select '选择第一行
        ' 循环
        For i = 1 To 10
            Selection.Copy
            ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.Select
            Selection.Insert Shift:=xlDown
        Next
    
    End Sub
    ```

  + :warning: 宏执行完不可撤销-->所以我们先将数据备份

    ![image-20200303103644076](Readme.assets/image-20200303103644076.png)











## 警告或者错误解决

### :warning: 文档检查器

![image-20200303095237074](Readme.assets/image-20200303095237074.png)

#### :ballot_box_with_check: 解决

[解决方案]( https://jingyan.baidu.com/article/7908e85cd88ce3af481ad2ad.html )



### :warning: 无法在未启用宏的工作簿中保存

![image-20200303132800111](Readme.assets/image-20200303132800111.png)

### :ballot_box_with_check: 解决

+ 点否

+ 出现保存对话框

  ![image-20200303132926259](Readme.assets/image-20200303132926259.png)

+ 选择启用宏的工作簿

  ![image-20200303132954163](Readme.assets/image-20200303132954163.png)



+ 出现感叹号

  ![image-20200303133021183](Readme.assets/image-20200303133021183.png)

+ 打开这个文件即可



### :warning:警告直接关闭 

> 发现下面代码未起作用

```vba
Application.DisplayAlerts = False '成对出现的

	'....code

Application.DisplayAlerts = True
```







## EXCEL常用



### :hammer: 菜单下拉

1. 选择需要有菜单下拉的单元格
2. 数据->数据验证->验证条件(允许:序列,来源:用逗号隔开选项)即可



### :hammer: 拖动列

`shift` + 鼠标



### :hammer: 自动换行

快捷键: `alt`+`enter`





