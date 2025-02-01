Attribute VB_Name = "tofliv"
Option Explicit

Public Function GetHdSn()
    Dim Locator As Object        'SWbemLocatorオブジェクト
    Dim Service As Object      'SWbemServicesExオブジェクト
    Dim ObjSet As Object      'SWbemObjectSetオブジェクト
    Dim ObjEx As Object        'SWbemObjectExオブジェクト
    Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Set Service = Locator.ConnectServer
    Set ObjSet = Service.ExecQuery("Select * From Win32_DiskDrive")
    For Each ObjEx In ObjSet
      If ObjEx.MediaType <> "Removable Media" Then
        GetHdSn = Trim(ObjEx.SerialNumber)
        Exit For
      End If
    Next
    Set ObjSet = Nothing
    Set Service = Nothing
    Set Locator = Nothing
End Function


Public Function Mitm(x_mitm_flag As Integer, y_mitm_flag As Integer, x_width As Double, y_width As Double)
    
    With ActiveChart
        If 1 = x_mitm_flag Then
            .Axes(xlCategory, 1).MinorTickMark = xlNone
            .Axes(xlCategory, 2).MinorTickMark = xlNone
        Else
            .Axes(xlCategory, 1).MinorTickMark = xlInside
            .Axes(xlCategory, 2).MinorTickMark = xlInside
            .Axes(xlCategory, 1).MinorUnit = x_width / x_mitm_flag
            .Axes(xlCategory, 2).MinorUnit = x_width / x_mitm_flag
        End If
        
        If 1 = y_mitm_flag Then
            .Axes(xlValue, 1).MinorTickMark = xlNone
            .Axes(xlValue, 2).MinorTickMark = xlNone
        Else
            .Axes(xlValue, 1).MinorTickMark = xlInside
            .Axes(xlValue, 2).MinorTickMark = xlInside
            .Axes(xlValue, 1).MinorUnit = y_width / y_mitm_flag
            .Axes(xlValue, 2).MinorUnit = y_width / y_mitm_flag
        End If
        
    End With
    
End Function


Public Function Tnf(XorY As Integer, tick_numfor As String, dp_g As Integer, width As Double)

    
    With ActiveChart.Axes(XorY, 1)
    
        If tick_numfor Like "*E+00" Then
            Select Case dp_g
                Case 1
                    .TickLabels.NumberFormatLocal = "0.E+00"
                Case 2
                    .TickLabels.NumberFormatLocal = "0.0E+00"
                Case 3
                    .TickLabels.NumberFormatLocal = "0.00E+00"
                Case 4
                    .TickLabels.NumberFormatLocal = "0.000E+00"
                Case 5
                    .TickLabels.NumberFormatLocal = "0.0000E+00"
                Case 6
                    .TickLabels.NumberFormatLocal = "0.00000E+00"
            End Select
        ElseIf Int(width) = width Then
            .TickLabels.NumberFormatLocal = "G/標準"
        ElseIf tick_numfor Like "*.*" Then
            Select Case dp_g
                Case 1
                    .TickLabels.NumberFormatLocal = "0."
                Case 2
                    .TickLabels.NumberFormatLocal = "0.0"
                Case 3
                    .TickLabels.NumberFormatLocal = "0.00"
                Case 4
                    .TickLabels.NumberFormatLocal = "0.000"
                Case 5
                    .TickLabels.NumberFormatLocal = "0.0000"
                Case 6
                    .TickLabels.NumberFormatLocal = "0.00000"
            End Select
            
        ElseIf tick_numfor Like "G/標準" Then
            If 1000 < width Then
                .TickLabels.NumberFormatLocal = "0.0E+00"
            ElseIf width Like "*.??" Then
                .TickLabels.NumberFormatLocal = "G/標準"
            ElseIf width Like "*.?" Then
                .TickLabels.NumberFormatLocal = "G/標準"
            ElseIf Int(width) = width Then
                 .TickLabels.NumberFormatLocal = "G/標準"
            Else
                .TickLabels.NumberFormatLocal = "0.0"
            End If
        End If
    
    End With

End Function



Public Function getTitf(XorY As Integer)


    With ActiveChart.Axes(XorY, 1)
        
        
        Dim getstr As String
        Dim t_count As Integer
        Dim a_flag As Boolean: a_flag = True
        
        t_count = Len(.AxisTitle.Characters.Text)
        
        Dim i As Integer
        For i = 1 To t_count
            Dim flag As Boolean
            Dim str As String
            flag = .AxisTitle.Characters(i, 1).Font.Italic
            str = .AxisTitle.Characters(i, 1).Text
            
            If flag And a_flag Then
                getstr = getstr + "\"
                getstr = getstr + str
                a_flag = False
            ElseIf flag Then
                getstr = getstr + str
            ElseIf Not flag And Not a_flag Then
                getstr = getstr + "\"
                getstr = getstr + str
                a_flag = True
            ElseIf Not flag And a_flag Then
                getstr = getstr + str
            End If
        Next
        
        If Len(getstr) - Len(Replace(getstr, "\", "")) = 1 Then
            getstr = getstr + "\"
        End If
        
    
    End With
    
     getTitf = getstr

End Function


Public Function setTitf(XorY As Integer, axis_title As String, axes_color As Long)

    With ActiveChart.Axes(XorY, 1)
    
        .HasTitle = True
        Dim title_fsize As Long
        title_fsize = .AxisTitle.Font.Size
        With .AxisTitle
            .Characters.Text = axis_title
            .Font.Color = axes_color
            If Not (InStrRev(axis_title, "\") = 0) Then
                .Characters.Text = axis_title
                On Error Resume Next
                    .Characters(1, InStr(axis_title, "\")).Font.Italic = False
                    .Characters(InStr(axis_title, "\"), _
                        InStrRev(axis_title, "\") - InStr(axis_title, "\")).Font.Italic = True
                    .Characters(InStrRev(axis_title, "\"), Len(axis_title)).Font.Italic = False
                .Characters(InStr(axis_title, "\"), 1).Delete
                .Characters(InStrRev(axis_title, "\") - 1, 1).Delete
                On Error GoTo 0
            Else
                .Font.Size = title_fsize
                .Font.Italic = False
            End If
        End With
        .TickLabels.Font.Color = axes_color
    End With

End Function


Public Function setLogMaxValue(XorY As Integer, log_base As Double)

    Application.ScreenUpdating = False
    
    On Error Resume Next
        Dim series_count As Long
        series_count = ActiveChart.SeriesCollection.Count
        If Err.Number <> 0 Then
            MsgBox "変更するグラフを選択してください。", vbOKOnly + vbCritical, "読み込みエラー"
            Exit Function
        End If
    On Error GoTo 0
    If ActiveChart.SeriesCollection(2).Name = "Non" Then
        series_count = 1
    End If

    Dim srcc As Integer
    For srcc = 1 To series_count
    
        Dim myarrx() As Variant
        myarrx = ActiveChart.SeriesCollection(srcc).XValues
        Dim flag As Boolean
        flag = True
        Dim ele_count As Integer
        Dim arrx(100000) As String
        ele_count = 1
        Do
            On Error Resume Next
                arrx(ele_count) = myarrx(ele_count)
                If Err.Number <> 0 Then
                    flag = False
                    ele_count = ele_count - 1
                Else
                    ele_count = ele_count + 1
                End If
            On Error GoTo 0
        Loop While flag
        
        
        
        Dim i As Integer
        
        Dim max_arr_x(1 To 100) As Double
        Dim min_arr_x(1 To 100) As Double
        For i = 1 To ele_count
            If Not (IsNumeric(arrx(i))) Then
            Else
                If 1 = i Then
                    max_arr_x(srcc) = arrx(i)
                    min_arr_x(srcc) = arrx(i)
                End If
                If max_arr_x(srcc) < arrx(i) Then
                    max_arr_x(srcc) = arrx(i)
                End If
                If min_arr_x(srcc) > arrx(i) Then
                    min_arr_x(srcc) = arrx(i)
                End If
            End If
        Next i
        
        Dim arr_y(100000) As String
        ActiveChart.SeriesCollection(srcc).HasDataLabels = True
        For i = 1 To ele_count
            arr_y(i) = ActiveChart.FullSeriesCollection(srcc).Points(i).DataLabel.Text
        Next i
        Dim max_arr_y(1 To 100) As Double
        Dim min_arr_y(1 To 100) As Double
        Dim ele_arr(1 To 100) As Double
        For i = 1 To ele_count
            If Not (IsNumeric(arr_y(i))) Then
            
            Else
                If 1 = i Then
                max_arr_y(srcc) = arr_y(i)
                min_arr_y(srcc) = arr_y(i)
                End If
                If max_arr_y(srcc) < arr_y(i) Then
                    max_arr_y(srcc) = arr_y(i)
                End If
                If min_arr_y(srcc) > arr_y(i) Then
                    min_arr_y(srcc) = arr_y(i)
                End If
            End If
            
        Next i
        ele_arr(srcc) = ele_count
        
        ActiveChart.SeriesCollection(srcc).HasDataLabels = False
        
    Next srcc
    
    Application.ScreenUpdating = True
    
    Dim max_x As Double, min_x As Double
    Dim max_ele As Double, min_ele As Double
    For i = 1 To series_count
        If 1 = i Then
            max_x = max_arr_x(i)
            min_x = min_arr_x(i)
            max_ele = ele_arr(i)
            min_ele = ele_arr(i)
        End If
        If max_x < max_arr_x(i) Then
            max_x = max_arr_x(i)
            max_ele = ele_arr(i)
        End If
        If min_x > min_arr_x(i) Then
            min_x = min_arr_x(i)
            min_ele = ele_arr(i)
        End If
    Next i
    
    Dim max_y As Double, min_y As Double
    For i = 1 To series_count
        If 1 = i Then
            max_y = max_arr_y(i)
            min_y = min_arr_y(i)
        End If
        If max_y < max_arr_y(i) Then
            max_y = max_arr_y(i)
        End If
        If min_y > min_arr_y(i) Then
            min_y = min_arr_y(i)
        End If
    Next i
    
    Dim max_t_x As Double, min_t_x As Double
    max_t_x = max_x * log_base
    Dim max_t_y As Double, min_t_y As Double
    max_t_y = max_y * log_base
    
    
    
    If 1 = XorY Then
        setLogMaxValue = max_t_x
    ElseIf 2 = XorY Then
        setLogMaxValue = max_t_y
    End If
    

End Function






Public Function setlegend()
        
    With ActiveChart
    
        .HasLegend = True
        .legend.Top = 10
        
    
        Dim series_count As Long
        series_count = .SeriesCollection.Count
        If 10 <= series_count Then
            MsgBox "系列が１０以上ある場合は自動設定できません。" _
                , vbCritical + vbQuestion, "警告"
            Exit Function
        End If
        
        Dim lf_size(1 To 20) As Double
        Dim lh(1 To 20) As Double
        Dim lw(1 To 20) As Double
        
        Dim i As Long
        For i = 1 To series_count
        
            lf_size(i) = .legend.LegendEntries(i).Format.TextFrame2.TextRange.Font.Size
            lh(i) = lf_size(i)
            lw(i) = 30 + (lf_size(i) * 0.25) * Len(.SeriesCollection(i).Name)
        
        Next i
                
        Dim lf_size_max As Double
        Dim lh_max As Double
        Dim lw_max As Double
        
        For i = 1 To series_count
        
            If 1 = i Then
                lf_size_max = lf_size(i)
                lh_max = lh(i)
                lw_max = lw(i)
            End If
            
            If lf_size_max < lf_size(i) Then
                lf_size_max = lf_size(i)
            End If
            If lh_max < lh(i) Then
                lh_max = lh(i)
            End If
            If lw_max < lw(i) Then
                lw_max = lw(i)
            End If
        
        Next i
        
        .legend.Height = lh_max * (series_count - 1)
        .legend.width = lw_max

        
        Dim one_src_flag As Boolean
        one_src_flag = False
        
        For i = 1 To series_count
            
            If "Non" = .SeriesCollection(i).Name Then
                ActiveChart.legend.LegendEntries(i).Delete
                one_src_flag = True
            End If
            
        Next i
        
        If one_src_flag And 2 = series_count Then
            .HasLegend = False
            Exit Function
        End If
        
        .legend.IncludeInLayout = False
        .legend.Format.Line.Visible = True
        .legend.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        
    End With



End Function



Public Function setSeriesMarkerFormat()

    Dim series_count As Long
    series_count = ActiveChart.SeriesCollection.Count
    If 10 <= series_count Then
        MsgBox "系列が１０以上ある場合は自動設定できません。" _
        , vbCritical + vbQuestion, "警告"
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To series_count
        Dim f As Boolean
        f = ActiveChart.SeriesCollection(i).Format.Line.Visible
        
        If f Then
            With ActiveChart.SeriesCollection(i).Format.Line
                .Weight = 0
                .ForeColor.RGB = RGB(0, 0, 0)
            End With
        End If
        
        With ActiveChart.SeriesCollection(i)
            .MarkerStyle = i
            .MarkerSize = 4
            .MarkerBackgroundColor = RGB(0, 0, 0)
            .MarkerForegroundColor = RGB(0, 0, 0)
        End With
    Next i

End Function



Public Function showProBar(str_cap As String, start_value As Double, ele_count As Long)
    
    If start_value < ele_count Then
        Load UserForm1
        UserForm1.Caption = str_cap
        UserForm1.Show vbModeless
    End If

End Function




Public Function getHttpAutho()

    Dim httpreq As Object
    Set httpreq = CreateObject("MSXML2.ServerXMLHTTP")
    httpreq.Open "GET", "https://sites.google.com/view/ryunens/認証ページ", False
    On Error Resume Next
    httpreq.send
    If Err.Number <> 0 Then
        getHttpAutho = "NetError"
        Exit Function
    End If
    On Error GoTo 0
    
    Do While httpreq.readyState < 4
        DoEvents
    Loop
    
    Dim htmlDoc As Object
    Set htmlDoc = New HTMLDocument
    htmlDoc.write httpreq.responseText
    If Not (httpreq.Status = "2??") Then
        getHttpAutho = httpreq.Status
    End If
    Dim re_https_status As String
    re_https_status = httpreq.statusText
    
    MsgBox "ネットワーク接続：" & re_https_status, vbOKOnly + vbInformation, "ネットワーク接続"
    DoEvents
    
    Dim autho_key_num(10) As String
    Dim i As Long: i = 0
    
    Dim objitem As Object
    For Each objitem In htmlDoc.getElementsByTagName("p")
        autho_key_num(i) = objitem.innerText
        i = i + 1
    Next
    
    getHttpAutho = autho_key_num(1)
    
    Set httpreq = Nothing
    Set htmlDoc = Nothing
    Set objitem = Nothing
    
End Function





















