'对象必须使用 set
'局部对象返回会导致销毁，外层拿不到数据

'全局变量
Public objSC, http, html, tmp
'获取网页内容
Public Function getHtml(url As String) As String
    Dim cookie As String
    Set http = CreateObject("Msxml2.ServerXMLHTTP")
    
    Debug.Print ("URL:" & url)
    If url <> "" Then
        cookie = localCache("get", "localCookie")
        http.Open "GET", url, False
        With http
            .setRequestHeader "Referer", url            '设置正确的Referer
            '.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .setRequestHeader "Connection", "keep-alive"
            If cookie <> "" Then
                .setRequestHeader "Cookie", cookie          '设置有效Cookie
            End If
        End With
        http.Send
        
        getHtml = http.responseText
        'ActiveSheet.Cells(5, 10) = getHtml
    Else
        getHtml = ""
    End If
End Function

'获取公司信息 天眼查版
Function getCompanyInfo(name As String) As Variant
    Dim text As String, obj, result, companyUrl As String, data(20, 3) As Variant, index As Integer, tmpIndex As Integer, cookie As String, cache As String
    '检查缓存
    cache = localCache("get", name)
    
    If cache = "" Then
        '搜索
        text = getHtml("https://www.tianyancha.com/search?checkFrom=searchBox&key=" & UrlEncode(name))
        toPast (text)
        '登录
        If InStr(text, "请输入您的手机号码") Then
            cookie = InputBox("已被网站屏蔽。打开浏览器，登录天眼查，使用开发人员工具获取cookie", "Cookie")
            cookie = localCache("set", "localCookie", cookie)
            Exit Function
        End If
        '您的访问过于频繁
        If InStr(text, "您的访问过于频繁") Then
            MsgBox ("您的访问过于频繁，被屏蔽")
            Exit Function
        End If
        '解析html
        Set html = CreateObject("htmlfile")
        html.designMode = "on" ' 开启编辑模式
        html.write text ' 写入数据
        
        Set obj = html.getElementsByTagName("a")      'getElementById getElementsByName
        
        For Each el In obj
          If el.innerText = name Then
            companyUrl = el.href
          End If
        Next
        'If commpanyUrl = "" Or Len(commpanyUrl) = 0 Then
        '    For Each el In obj
        '        If InStr(el.href, "www.tianyancha.com/company/") Then
        '            MsgBox ("企业可能已更名")
        '            companyUrl = el.href
        '        End If
        '    Next
        'End If
        
        If companyUrl <> "" And TypeName(companyUrl) <> "Nothing" And Len(companyUrl) <> 0 Then
            '详情
            text = getHtml(companyUrl)
            toPast (text)
            Set html = CreateObject("htmlfile")
            html.designMode = "on" '
            html.write text '
            Set obj = html.getElementById("_container_baseInfo").getElementsByTagName("td")
            index = 0
            For Each el In obj
                '注册资本
                If InStr(el.innerText, "注册资本") Then
                    For Each ell In el.Children
                        tmpIndex = 0
                        For Each elll In ell.Children
                            data(index, tmpIndex) = elll.innerText
                            tmpIndex = tmpIndex + 1
                        Next
                        index = index + 1
                    Next
                '其他信息
                Else
                    If UBound(Split(el.innerText, "：")) >= 1 Then
                        data(index, 0) = Split(el.innerText, "：")(0)
                        data(index, 1) = Split(el.innerText, "：")(1)
                        index = index + 1
                    End If
                    
                End If
            Next
            '基础信息
            'Set obj = html.getElementById("company_web_top")
            'MsgBox (obj.class)
            
        End If
        
        getCompanyInfo = data
        '缓存
        tmp = localCache("set", name, dumpArray(data))
    Else
        getCompanyInfo = parseArray(cache)
    End If
End Function
Sub test()
    Debug.Print (arrayToString(getCompanyInfoQCC("上海聚洋国际货物运输代理有限公司")))
    'Debug.Print localCache("del", "ddd")
    'Debug.Print (Tyc("宝号酒店管理咨询（上海）有限公司", "注册资本"))
   
End Sub
'缓存选中公司数据
Sub cacheCompayInfo()
    'Dim name As String
    'name = ActiveCell.value

    Dim rng As Range
    On Error Resume Next
    Set rng = Selection                 'rng是一个range对象，希望能得到该rng就是选择的区域，按这样写法是错的，请问如何写
    If rng Is Nothing Then              '想通过判断如果，没有选择到区域的话，就提示，但rng=nothing的写法也是错的，请问如何写
        MsgBox ("请选择需要缓存的公司！")
    End If
    For Each x In rng
        If x <> "" And TypeName(x) <> "Nothing" And Len(x) <> 0 Then
            getCompanyInfoQCC (x)
        End If
    Next
End Sub
'获取公司信息 企查查版
Function getCompanyInfoQCC(name As String) As Variant
    Dim text As String, obj, result, companyUrl As String, data(30, 3) As Variant, index As Integer, tmpIndex As Integer, cookie As String, cache As String
    '检查缓存
    cache = localCache("get", name)
    
    If cache = "" Then
        '搜索
        text = getHtml("https://www.qichacha.com/search?key=" & UrlEncode(name))
        toPast (text)
        '登录
        If InStr(text, "请输入您的手机号码") Then
            'cookie = InputBox("已被网站屏蔽。打开浏览器，登录天眼查，使用开发人员工具获取cookie", "Cookie")
            'cookie = localCache("set", "localCookie", cookie)
            'Exit Function
        End If
        '您的访问过于频繁
        If InStr(text, "您的访问过于频繁") Then
            MsgBox ("您的访问过于频繁，被屏蔽")
            Exit Function
        End If
        '解析html
        Set html = CreateObject("htmlfile")
        html.designMode = "on" ' 开启编辑模式
        html.write text ' 写入数据
        
        Set obj = html.getElementsByTagName("a")      'getElementById getElementsByName
        
        For Each el In obj
          'Debug.Print (el.innerText)
          If el.innerText = name Then
            companyUrl = el.href
          End If
        Next
        
        If companyUrl <> "" And TypeName(companyUrl) <> "Nothing" And Len(companyUrl) <> 0 Then
            '详情
            Debug.Print (companyUrl)
            text = getHtml("https://www.qichacha.com/" & Replace(companyUrl, "about:/", ""))
            toPast (text)
            Set html = CreateObject("htmlfile")
            html.designMode = "on" '
            html.write text
            
            data(0, 0) = "公司名称"
            data(0, 1) = name
            index = 1
            '基础信息
            Set obj = html.getElementById("company-top").getElementsByTagName("small")
            For Each el In obj
                Dim tag As Boolean
                tag = False
                For Each x In Split(el.innerText, " ")
                    'Debug.Print (x)
                    If tag Then
                        data(index, 1) = x
                        index = index + 1
                        tag = False
                    End If
                    If InStr(x, "电话") Then
                        data(index, 0) = "电话"
                        data(index, 1) = Replace(x, "电话：", "")
                        index = index + 1
                    End If
                    If InStr(x, "官网") Then
                        data(index, 0) = "官网"
                        data(index, 1) = Replace(x, "官网：", "")
                        index = index + 1
                    End If
                    If InStr(x, "邮箱") Then
                        data(index, 0) = "邮箱"
                        tag = True
                    End If
                    If InStr(x, "地址") Then
                        data(index, 0) = "地址"
                        tag = True
                    End If
                Next
            Next
            
        
            
            '基本信息
            If Not InStr(text, "统一社会信用代码") Then
                text = getHtml("https://www.qichacha.com/company_getinfos?tab=base&unique=" & Replace(Replace(companyUrl, "about:/firm_", ""), ".html", "") & "&companyname=" & UrlEncode(name))
                toPast (text)
                Set html = CreateObject("htmlfile")
                html.designMode = "on" '
                html.write text
            End If
            Set obj = html.getElementsByTagName("table")
            For Each el In obj
                'Debug.Print (el.innerHtml)
                If InStr(el.innerText, "统一社会信用代码") Then
                    For Each etr In el.getElementsByTagName("tr")
                        tmpIndex = 0
                        For Each ell In etr.Children
                            data(index, tmpIndex) = Replace(Replace(Trim(ell.innerText), "：", ""), "【依法须经批准的项目，经相关部门批准后方可开展经营活动】", "")
                            If (tmpIndex = 1) Then
                                index = index + 1
                                tmpIndex = 0
                            Else
                                tmpIndex = tmpIndex + 1
                            End If
                        Next
                    Next
                End If
                
            Next
            
        End If
        getCompanyInfoQCC = data
        '缓存
        tmp = localCache("set", name, dumpArray(data))
    Else
        getCompanyInfoQCC = parseArray(cache)
    End If
End Function

'自动填充选中单元格
Sub writeCompayInfo()
    'Dim name As String
    'name = ActiveCell.value

    Dim rng As Range
    On Error Resume Next
    Set rng = Selection                 'rng是一个range对象，希望能得到该rng就是选择的区域，按这样写法是错的，请问如何写
    If rng Is Nothing Then              '想通过判断如果，没有选择到区域的话，就提示，但rng=nothing的写法也是错的，请问如何写
        MsgBox ("请选择需要缓存的公司！")
    End If
    Dim title As String, oldValue As String
    For Each x In rng
        If x <> "" And TypeName(x) <> "Nothing" And Len(x) <> 0 Then
            Debug.Print ("" & x.row & ":" & x.Column)
            For i = 1 To 30
                title = ActiveSheet.Cells(1, i)
                oldValue = ActiveSheet.Cells(x.row, i)
                If title <> "" And Len(title) <> 0 And i <> x.Column And (oldValue = "" Or Len(oldValue) = 0 Or oldValue = "N/A") Then
                    'Debug.Print ActiveSheet.Cells(1, i)
                    ActiveSheet.Cells(x.row, i) = Qcc(x.value, title)
                End If
            Next
        End If
    Next
End Sub
'获取企业的某项信息 企查查
Public Function Qcc(name As String, field As String, Optional field2 As String = "", Optional field3 As String = "") As String
    Dim data As Variant
    Qcc = ""
    data = getCompanyInfoQCC(name)
    For i = LBound(data, 1) To UBound(data, 1)
        If data(i, 0) = field Then
            Qcc = data(i, 1)
        End If
    Next
    If Qcc = "" And field2 <> "" Then
        For i = LBound(data, 1) To UBound(data, 1)
            If data(i, 0) = field2 Then
                Qcc = data(i, 1)
            End If
        Next
    End If
    If Qcc = "" And field3 <> "" Then
        For i = LBound(data, 1) To UBound(data, 1)
            If data(i, 0) = field3 Then
                Qcc = data(i, 1)
            End If
        Next
    End If
    If Qcc = "" Then
        Qcc = "N/A"
    End If
End Function
'获取企业的某项信息 天眼查
Public Function Tyc(name As String, field As String)
    Dim data As Variant
    Tyc = "N/A"
    data = getCompanyInfo(name)
    For i = LBound(data, 1) To UBound(data, 1)
        If data(i, 0) = field Then
            Tyc = data(i, 1)
        End If
    Next
End Function
'字符串转数组
Public Function stringToArray(str As String) As Object
    Dim strJSON
    Set objSC = CreateObject("ScriptControl")    '调用ScriptControl对象
    strJSON = "var o=" & str & ";"
    objSC.Language = "JScript"
    objSC.AddCode (strJSON)
     
    Set stringToArray = objSC.CodeObject.o
    'CallByName(stringToArray, "myname", VbGet)
End Function
'数据转json字符串
Public Function arrayToString(arr As Variant) As String
    arrayToString = arrayToString & "["
    For i = LBound(arr, 1) To UBound(arr, 1)
        arrayToString = arrayToString & "["
        For ii = LBound(arr, 2) To UBound(arr, 2)
            arrayToString = arrayToString & Chr(34) & arr(i, ii) & Chr(34) & ","
        Next
        arrayToString = Left(arrayToString, Len(arrayToString) - 1)
        arrayToString = arrayToString & "],"
    Next
    
    arrayToString = Left(arrayToString, Len(arrayToString) - 1)
    arrayToString = arrayToString & "]"
End Function
'数组转字符串 自定义
Public Function dumpArray(arr As Variant) As String
    For i = LBound(arr, 1) To UBound(arr, 1)
        For ii = LBound(arr, 2) To UBound(arr, 2)
            dumpArray = dumpArray & arr(i, ii) & "#"
        Next
        dumpArray = Left(dumpArray, Len(dumpArray) - 1)
        dumpArray = dumpArray & "$"
    Next
    
    dumpArray = Left(dumpArray, Len(dumpArray) - 1)
End Function
'字符串转数组 自定义
Public Function parseArray(str As String) As Variant
    Dim row, col, r, c
    row = UBound(Split(str, "$"))
    col = UBound(Split(Split(str, "$")(0), "#"))
    ReDim data(row, col) As Variant
    r = 0
    For Each row In Split(str, "$")
        c = 0
        For Each cell In Split(row, "#")
            data(r, c) = cell
            c = c + 1
        Next
        r = r + 1
    Next
    parseArray = data
End Function
'本地缓存操作
Public Function localCache(opt As String, key As String, Optional value As String = "") As String
    Dim sheet As Worksheet, hasSheet As Boolean, currentSheet, rows, tag, dataSheet
    hasSheet = False
    currentSheet = ActiveSheet.name
    
    For Each sheet In Worksheets
        If sheet.name = "系统缓存数据" Then
            'Sheets(sheet.name).Select
            hasSheet = True
        End If
    Next
    
    If hasSheet = False Then
        Worksheets.Add
        ActiveSheet.name = "系统缓存数据"
    End If
    
    Set dataSheet = Sheets("系统缓存数据")
    
    rows = dataSheet.Cells(dataSheet.rows.Count, 1).End(xlUp).row
    
    oldKeys = dataSheet.Range("a1:b" & rows)
    
    '保存
    If opt = "set" Then
        tag = False
        For i = 1 To UBound(oldKeys, 1)
            If oldKeys(i, 1) = key Then
                tag = True
                dataSheet.Cells(i, 2) = value
            End If
        Next
        If Not tag Then
            dataSheet.Cells(rows + 1, 1) = key
            dataSheet.Cells(rows + 1, 2) = value
        End If
    End If
    '取值
    If opt = "get" Then
        tag = False
        For i = 1 To UBound(oldKeys, 1)
            If oldKeys(i, 1) = key Then
                tag = True
                value = oldKeys(i, 2)
            End If
        Next
    End If
    '删除
    If opt = "del" Or opt = "delete" Then
        tag = False
        For i = 1 To UBound(oldKeys, 1)
            If oldKeys(i, 1) = key Then
                tag = True
                value = oldKeys(i, 2)
                dataSheet.Cells(i, 1).MergeArea.EntireRow.Delete
            End If
        Next
    End If
    
    Sheets(currentSheet).Select
    
    localCache = value
End Function

'放置到剪切板
Public Function toPast(text As String)
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")    '得到的字符串放入剪贴板，记事本观察数据
        .SetText text                                                  '数据正常显示，可以提取了
        .PutInClipboard
    End With
End Function
'根据id获取元素
Public Function getElementById(id As String, text As String) As Object
    Set html = CreateObject("htmlfile")
    html.designMode = "on" ' 开启编辑模式
    html.write text ' 写入数据
    Set getElementById = html.getElementById(id)
End Function
'打印数组
Public Function printArray(arr) As String
    Dim row
    Debug.Print ("行：" & UBound(arr, 1))
    Debug.Print ("列：" & UBound(arr, 2))
    For i = LBound(arr, 1) To UBound(arr, 1)
        row = ""
        For ii = LBound(arr, 2) To UBound(arr, 2)
            row = row & arr(i, ii) & "|"
        Next
        printArray = printArray & row
        Debug.Print (row)
    Next
End Function
Function URLDecode(ByVal What)
'URL decode Function
'2001 Antonin Foller, PSTRUH Software, http://www.motobit.com
    Dim Pos, pPos
    
    'replace + To Space
    What = Replace(What, "+", " ")
    
    On Error Resume Next
    Dim Stream: Set Stream = CreateObject("ADODB.Stream")
    If Err = 0 Then 'URLDecode using ADODB.Stream, If possible
        On Error GoTo 0
        Stream.Type = 2 'String
        Stream.Open
        
        'replace all %XX To character
        Pos = InStr(1, What, "%")
        pPos = 1
        Do While Pos > 0
            Stream.WriteText Mid(What, pPos, Pos - pPos) + _
            Chr(CLng("&H" & Mid(What, Pos + 1, 2)))
            pPos = Pos + 3
            Pos = InStr(pPos, What, "%")
        Loop
        Stream.WriteText Mid(What, pPos)
    
        'Read the text stream
        Stream.Position = 0
        URLDecode = Stream.ReadText
    
        'Free resources
        Stream.Close
    Else 'URL decode using string concentation
        On Error GoTo 0
        'UfUf, this is a little slow method.
        'Do Not use it For data length over 100k
        Pos = InStr(1, What, "%")
        Do While Pos > 0
            What = Left(What, Pos - 1) + _
            Chr(CLng("&H" & Mid(What, Pos + 1, 2))) + _
            Mid(What, Pos + 3)
            Pos = InStr(Pos + 1, What, "%")
        Loop
        URLDecode = What
    End If
End Function

Public Function UrlEncode(ByRef szString As String) As String
    Dim szChar   As String
    Dim szTemp   As String
    Dim szCode   As String
    Dim szHex    As String
    Dim szBin    As String
    Dim iCount1  As Integer
    Dim iCount2  As Integer
    Dim iStrLen1 As Integer
    Dim iStrLen2 As Integer
    Dim lResult  As Long
    Dim lAscVal  As Long
    szString = Trim$(szString)
    iStrLen1 = Len(szString)
    For iCount1 = 1 To iStrLen1
        szChar = Mid$(szString, iCount1, 1)
        lAscVal = AscW(szChar)
        If lAscVal >= &H0 And lAscVal <= &HFF Then
            If (lAscVal >= &H30 And lAscVal <= &H39) Or _
            (lAscVal >= &H41 And lAscVal <= &H5A) Or _
            (lAscVal >= &H61 And lAscVal <= &H7A) Then
                szCode = szCode & szChar
            Else
                szCode = szCode & "%" & Hex(AscW(szChar))
            End If
        Else
            szHex = Hex(AscW(szChar))
            iStrLen2 = Len(szHex)
            For iCount2 = 1 To iStrLen2
                szChar = Mid$(szHex, iCount2, 1)
                Select Case szChar
                    Case Is = "0"
                    szBin = szBin & "0000"
                    Case Is = "1"
                    szBin = szBin & "0001"
                    Case Is = "2"
                    szBin = szBin & "0010"
                    Case Is = "3"
                    szBin = szBin & "0011"
                    Case Is = "4"
                    szBin = szBin & "0100"
                    Case Is = "5"
                    szBin = szBin & "0101"
                    Case Is = "6"
                    szBin = szBin & "0110"
                    Case Is = "7"
                    szBin = szBin & "0111"
                    Case Is = "8"
                    szBin = szBin & "1000"
                    Case Is = "9"
                    szBin = szBin & "1001"
                    Case Is = "A"
                    szBin = szBin & "1010"
                    Case Is = "B"
                    szBin = szBin & "1011"
                    Case Is = "C"
                    szBin = szBin & "1100"
                    Case Is = "D"
                    szBin = szBin & "1101"
                    Case Is = "E"
                    szBin = szBin & "1110"
                    Case Is = "F"
                    szBin = szBin & "1111"
                    Case Else
                End Select
            Next iCount2
            szTemp = "1110" & Left$(szBin, 4) & "10" & Mid$(szBin, 5, 6) & "10" & Right$(szBin, 6)
            For iCount2 = 1 To 24
                If Mid$(szTemp, iCount2, 1) = "1" Then
                    lResult = lResult + 1 * 2 ^ (24 - iCount2)
                Else: lResult = lResult + 0 * 2 ^ (24 - iCount2)
                End If
            Next iCount2
            szTemp = Hex(lResult)
            szCode = szCode & "%" & Left$(szTemp, 2) & "%" & Mid$(szTemp, 3, 2) & "%" & Right$(szTemp, 2)
        End If
        szBin = vbNullString
        lResult = 0
    Next iCount1
    UrlEncode = szCode
End Function


