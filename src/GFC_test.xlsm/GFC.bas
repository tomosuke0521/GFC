Attribute VB_Name = "GFC"
Option Explicit


Sub GraphSeriesAuto()
Attribute GraphSeriesAuto.VB_ProcData.VB_Invoke_Func = "G\n14"
    'Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("オプション")
        Dim xdp_g As Integer
        xdp_g = .Range("H9").Value
        Dim ydp_g As Integer
        ydp_g = .Range("H10").Value
        Dim x_mitm_flag As Integer
        x_mitm_flag = .Range("H13").Value
        Dim y_mitm_flag As Integer
        y_mitm_flag = .Range("H14").Value
        Dim rgb_r, rgb_g, rgb_b As Integer
        rgb_r = .Range("H17").Value - 1
        rgb_g = .Range("H22").Value - 1
        rgb_b = .Range("H27").Value - 1
    End With
    
    Dim HD_S_number As String
    HD_S_number = ThisWorkbook.Sheets("ホーム").Cells(1, 1).Value
    
    
    'Application.ScreenUpdating = True
    

    If tofliv.GetHdSn() = HD_S_number Then
    
        '情報取得セクション
        On Error Resume Next
            Dim series_count As Long
            series_count = ActiveChart.SeriesCollection.Count
            If Err.Number <> 0 Then
                MsgBox "変更するグラフを選択してください。", vbOKOnly + vbCritical, "読み込みエラー"
                Exit Sub
            End If
        On Error GoTo 0
        
        Dim Is_oneSeries As Boolean
            Is_oneSeries = True
        If 2 <= series_count Then
            With ActiveChart
                .SeriesCollection(series_count).AxisGroup = 2
                .HasAxis(xlCategory, 2) = True
            End With
        Else
            On Error Resume Next
                With ActiveChart
                    .SeriesCollection.NewSeries.Name = "Non"
                    .SeriesCollection(2).AxisGroup = 2
                    .HasAxis(xlCategory, 2) = True
                End With
                Is_oneSeries = False
                If Err.Number <> 0 Then
                    MsgBox "対応しているのは、「散布図」のみです", vbCritical
                    Exit Sub
                End If
            On Error GoTo 0
        End If
        
        
        Dim gatv As getautotickvalue
        Set gatv = New getautotickvalue
        With gatv
            Call .GetAutoMaxMinTickValue(Is_oneSeries)
            Dim x_max As Double: x_max = .xMaxAutoTickValue
            Dim x_min As Double: x_min = .xMinAutoTickValue
            Dim y_max As Double: y_max = .yMaxAutoTickValue
            Dim y_min As Double: y_min = .yMinAutoTickValue
        End With
        '軸のタイトルの取得
        On Error Resume Next
            Dim get_x_axis_title As String
            get_x_axis_title = tofliv.getTitf(1)
            Dim get_y_axis_title As String
            get_y_axis_title = tofliv.getTitf(2)
        On Error GoTo 0

        Dim get_x_width As Double
        get_x_width = (x_max - x_min) / ActiveChart.Axes(xlCategory, 1).MajorUnit - 1
        Dim get_y_width As Double
        get_y_width = (y_max - y_min) / ActiveChart.Axes(xlValue, 1).MajorUnit - 1
        
        
        Dim x_width_box As New ilintickvalbox
        Call x_width_box.GetInf("ｘ軸：目盛り数選択", "ｘ", get_x_width)
        Dim inp_x_width As String
        inp_x_width = x_width_box.UserInput
        Dim x_width As Double
        x_width = (x_max - x_min) / (inp_x_width + 1)
        
        Dim y_width_box As New ilintickvalbox
        Call y_width_box.GetInf("ｙ軸：目盛り数選択", "ｙ", get_y_width)
        Dim inp_y_width As Double
        inp_y_width = y_width_box.UserInput
        Dim y_width As Double
        y_width = (y_max - y_min) / (inp_y_width + 1)
        
        Dim c_x_axis_title As New ititlenamebox
        Call c_x_axis_title.GetInf("ｘ軸：タイトル設定", "ｘ", get_x_axis_title)
        Dim x_axis_title As String
        x_axis_title = c_x_axis_title.UserInput
        
        Dim c_y_axis_title As New ititlenamebox
        Call c_y_axis_title.GetInf("ｙ軸：タイトル設定", "ｙ", get_y_axis_title)
        Dim y_axis_title As String
        y_axis_title = c_y_axis_title.UserInput
        
        
        
'------------書式設定セクション------------

        Dim inp_axes_color As Long
        inp_axes_color = RGB(rgb_r, rgb_g, rgb_b)
        
        Call tofliv.Mitm(x_mitm_flag, y_mitm_flag, x_width, y_width)
        
        With ActiveChart
        
            '第１ｘ軸xlCategory
            With .Axes(xlCategory, 1)
                .MajorTickMark = xlInside
                .Format.Line.ForeColor.RGB = inp_axes_color
                On Error Resume Next
                    .MaximumScale = x_max
                    .MinimumScale = x_min
                    
                    .MajorUnit = x_width
                     If Err.Number <> 0 Then
                        MsgBox "申し訳ありませんが、対数軸には対応しておりません。", vbOKOnly + vbCritical, "読み込みエラー"
                        Exit Sub
                     End If
                On Error GoTo 0
                    
                Dim xtick_numfor As String
                xtick_numfor = ActiveChart.Axes(xlCategory, 1).TickLabels.NumberFormatLocal
                Call tofliv.Tnf(1, xtick_numfor, xdp_g, x_width)
                
                Call tofliv.setTitf(1, x_axis_title, inp_axes_color)
                
                .TickLabels.Font.Color = inp_axes_color
            End With
            
            '第１ｙ軸
            With .Axes(xlValue, 1)
                .Format.Line.ForeColor.RGB = inp_axes_color
                On Error Resume Next
                    .MaximumScale = y_max
                    .MinimumScale = y_min
                    
                    .MajorUnit = y_width
                    If Err.Number <> 0 Then
                        MsgBox "申し訳ありませんが、対数軸には対応しておりません。", vbOKOnly + vbCritical, "読み込みエラー"
                        Exit Sub
                    End If
                On Error GoTo 0
                
                Dim ytick_numfor As String
                ytick_numfor = ActiveChart.Axes(xlValue, 1).TickLabels.NumberFormatLocal
                Call tofliv.Tnf(2, ytick_numfor, ydp_g, y_width)
                
                Call tofliv.setTitf(2, y_axis_title, inp_axes_color)
                
                .TickLabels.Font.Color = inp_axes_color
            End With
            
            '第２ｘ軸
            With .Axes(xlCategory, 2)
                .TickLabelPosition = xlTickLabelPositionNone
                .MajorTickMark = xlInside
                On Error Resume Next
                    .MaximumScale = x_max
                    .MinimumScale = x_min
                    .MajorUnit = x_width
                    If Err.Number <> 0 Then
                        MsgBox "申し訳ありませんが、対数軸には対応しておりません。", vbOKOnly + vbCritical, "読み込みエラー"
                        Exit Sub
                    End If
                On Error GoTo 0
                
                .Format.Line.ForeColor.RGB = inp_axes_color
            End With
            
            '第２ｙ軸
            With .Axes(xlValue, 2)
                .TickLabelPosition = xlTickLabelPositionNone
                .TickLabelPosition = xlTickLabelPositionNone
                .MajorTickMark = xlInside
                On Error Resume Next
                    .MaximumScale = y_max
                    .MinimumScale = y_min
                    .MajorUnit = y_width
                     If Err.Number <> 0 Then
                        MsgBox "申し訳ありませんが、対数軸には対応しておりません。", vbOKOnly + vbCritical, "読み込みエラー"
                        Exit Sub
                    End If
                On Error GoTo 0
                
                .Format.Line.ForeColor.RGB = inp_axes_color
            End With
            
            'グラフ全体の書式設定
        
            .ChartArea.Format.Line.Visible = msoFalse
            
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).Crosses = xlMinimum
            
            .Axes(xlValue).HasMajorGridlines = False
            .Axes(xlValue).Crosses = xlMinimum
            
            Dim title_flag As Boolean
            title_flag = ActiveChart.HasTitle
            .HasTitle = title_flag
            
            Dim legend_flag As Boolean
            legend_flag = .HasLegend
            .HasLegend = legend_flag
            If legend_flag Then
                
            End If
            
            If 1 = ThisWorkbook.Sheets("オプション").Range("M12").Value Then
                Call tofliv.setlegend
            End If
            
            If 1 = ThisWorkbook.Sheets("オプション").Range("M16").Value Then
                Call tofliv.setSeriesMarkerFormat
            End If
            
        End With
    
    Else
    
        MsgBox "申し訳ありませんが、このコンピューターでは使えません。" _
        & vbCrLf & _
        "ご購入の検討をお願い申し上げます。", vbCritical + vbQuestion, "警告"
    
    End If
    
    
End Sub





'-----------------------------------ログバージョン-----------------------------------



Sub GraphSeriesLogAuto()
Attribute GraphSeriesLogAuto.VB_ProcData.VB_Invoke_Func = "A\n14"


    With ThisWorkbook.Sheets("オプション")
        Dim xdp_g As Integer
        xdp_g = .Range("H9").Value
        Dim ydp_g As Integer
        ydp_g = .Range("H10").Value
        Dim x_mitm_flag As Integer
        x_mitm_flag = .Range("H13").Value
        Dim y_mitm_flag As Integer
        y_mitm_flag = .Range("H14").Value
        Dim rgb_r, rgb_g, rgb_b As Integer
        rgb_r = .Range("H17").Value - 1
        rgb_g = .Range("H22").Value - 1
        rgb_b = .Range("H27").Value - 1
    End With
    
    If tofliv.GetHdSn() = ThisWorkbook.Sheets("ホーム").Cells(1, 1).Value Then
        
         '情報取得セクション
        On Error Resume Next
            Dim series_count As Long
            series_count = ActiveChart.SeriesCollection.Count
            If Err.Number <> 0 Then
                MsgBox "変更するグラフを選択してください。", vbOKOnly + vbCritical, "読み込みエラー"
                Exit Sub
            End If
        On Error GoTo 0
        
        
        '系列の個数の判定と系列の追加
        Dim auto_flag As Boolean
            auto_flag = True
        If 2 <= series_count Then
            With ActiveChart
                .SeriesCollection(series_count).AxisGroup = 2
                .HasAxis(xlCategory, 2) = True
            End With
        Else
            With ActiveChart
                .SeriesCollection.NewSeries.Name = "Non"
                .SeriesCollection(2).AxisGroup = 2
                .HasAxis(xlCategory, 2) = True
            End With
            auto_flag = False
        End If
        
        
        '軸の最小値と最大値の取得
        Dim x1_max As Double, x1_min As Double
        Dim y1_max As Double, y1_min As Double
        Dim x2_max As Double, x2_min As Double
        Dim y2_max As Double, y2_min As Double
        Dim x_max As Double, x_min As Double
        Dim y_max As Double, y_min As Double
        
        x1_max = ActiveChart.Axes(xlCategory, 1).MaximumScale
        Dim gtv As gettickvalue
        Set gtv = New gettickvalue
        x1_min = gtv.xMinTickValue
        y1_max = ActiveChart.Axes(xlValue, 1).MaximumScale
        y1_min = gtv.yMinTickValue
        x2_max = ActiveChart.Axes(xlCategory, 2).MaximumScale
        x2_min = ActiveChart.Axes(xlCategory, 2).MinimumScale
        y2_max = ActiveChart.Axes(xlValue, 2).MaximumScale
        y2_min = ActiveChart.Axes(xlValue, 2).MinimumScale
        
        Dim inp_axes_color As Long
        inp_axes_color = RGB(0, 0, 0)
        
        On Error Resume Next
            Dim get_x_axis_title As String
            get_x_axis_title = tofliv.getTitf(1)
            Dim get_y_axis_title As String
            get_y_axis_title = tofliv.getTitf(2)
        On Error GoTo 0
        
        Dim c_x_axis_title As New ititlenamebox
        Call c_x_axis_title.GetInf("ｘ軸：タイトル設定", "ｘ", get_x_axis_title)
        Dim x_axis_title As String
        x_axis_title = c_x_axis_title.UserInput
        Dim c_y_axis_title As New ititlenamebox
        Call c_y_axis_title.GetInf("ｙ軸：タイトル設定", "ｙ", get_y_axis_title)
        Dim y_axis_title As String
        y_axis_title = c_y_axis_title.UserInput
        
        
        Dim sele_log_axis As Long
        
        If MsgBox("ｘ軸を対数表示にしますか？" & vbCrLf & "「いいえ」の場合には線形軸にします。", vbYesNo, "ｘ軸選択") = vbYes Then
        
            If MsgBox("ｙ軸を対数表示にしますか？" & vbCrLf & "「いいえ」の場合には線形軸にします。", vbYesNo, "ｙ軸選択") = vbYes Then
                
                sele_log_axis = 1
            
            Else
                
                sele_log_axis = 2
            
            End If
                
        Else
            If MsgBox("ｙ軸を対数表示にしますか？" & vbCrLf & "「いいえ」の場合には線形軸にします。", vbYesNo, "ｙ軸選択") = vbYes Then
                
                sele_log_axis = 3
            Else
            
                sele_log_axis = 4
                
            End If
        End If
        
        Select Case sele_log_axis
            
            Case 1 'ｘ軸対数ｙ軸対数
            
                Dim p_x_inp_box As New ilogbasebox
                Call p_x_inp_box.GetInf("ｘ軸：基数設定", "x", 10)
                Dim p_x_log_base As Double
                p_x_log_base = p_x_inp_box.UserInput
                
                Dim p_y_inp_box As New ilogbasebox
                Call p_y_inp_box.GetInf("ｙ軸：基数設定", "ｙ", 10)
                Dim p_y_log_base As Double
                p_y_log_base = p_y_inp_box.UserInput
                
                x1_min = gtv.xLogMinTickValue(p_x_log_base)
                    
                If 0 < x1_min Then
                    x2_min = x1_min
                    x2_max = x1_max
                Else
                    MsgBox "対数軸に0以下の値は用いられません。" _
                    & vbCrLf & _
                    "軸の最小値が正の数になっている事を確認してください。", vbCritical + vbQuestion, "警告"
                    
                    Dim p_x_inp_min_box As New ilogminvalbox
                    Call p_x_inp_min_box.GetInf("ｘ軸：最小値設定", "ｘ", x1_max, 1)
                    
                    x1_min = p_x_inp_min_box.UserInput
                    x2_min = x1_min
                    x1_max = tofliv.setLogMaxValue(1, p_x_log_base)
                    x2_max = x1_max
                End If
                
                y1_min = gtv.yLogMinTickValue(p_y_log_base)
                
                If 0 < y1_min Then
                    y2_min = y1_min
                    y2_max = y1_max
                Else
                    MsgBox "対数軸に負の値は用いられません。" _
                    & vbCrLf & _
                    "軸の値が正の数になっている事を確認してください。", vbCritical + vbQuestion, "警告"
                    Dim p_y_inp_min_box As New ilogminvalbox
                    Call p_y_inp_min_box.GetInf("ｙ軸：最小値設定", "ｙ", y1_max, 1)
                    y1_min = p_y_inp_min_box.UserInput
                    y2_min = y1_min
                    y1_max = tofliv.setLogMaxValue(2, p_y_log_base)
                    y2_max = y1_max
                End If
                
                
                'グラフ書式設定
                With ActiveChart
                    .SetElement msoElementPrimaryCategoryAxisLogScale
                    .SetElement msoElementSecondaryCategoryAxisLogScale
                    .SetElement msoElementPrimaryValueAxisLogScale
                    .SetElement msoElementSecondaryValueAxisLogScale
                    
                    With .Axes(xlCategory, 1)
                        .MaximumScale = x1_max
                        .MinimumScale = x1_min
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabels.Font.Color = inp_axes_color
                        .HasTitle = True
                        Call tofliv.setTitf(1, x_axis_title, inp_axes_color)
                    End With
                    
                    With .Axes(xlCategory, 2)
                        .MaximumScale = x2_max
                        .MinimumScale = x2_min
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabelPosition = xlTickLabelPositionNone
                    End With
                    
                    With .Axes(xlValue, 1)
                        .MaximumScale = y1_max
                        .MinimumScale = y1_min
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabels.Font.Color = inp_axes_color
                        .HasTitle = True
                        Call tofliv.setTitf(2, y_axis_title, inp_axes_color)
                    End With
                    
                    With .Axes(xlValue, 2)
                        .MaximumScale = y2_max
                        .MinimumScale = y2_min
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabelPosition = xlTickLabelPositionNone
                    End With
                    
                    
                    .Axes(xlCategory).HasMajorGridlines = False
                    .Axes(xlCategory).Crosses = xlMinimum
                    .Axes(xlValue).HasMajorGridlines = False
                    .Axes(xlValue).Crosses = xlMinimum
                    
                End With
                
                
                
            
            Case 2 'ｘ軸対数ｙ軸線形
            
                Dim s_inp_box As New ilogbasebox
                Call s_inp_box.GetInf("ｘ軸：基数設定", "x", 10)
                Dim s_x_log_base As Double
                s_x_log_base = s_inp_box.UserInput
                    
                If 0 < x1_min Then
                    x2_min = x1_min
                    x2_max = x1_max
                Else
                    MsgBox "対数軸に0以下の値は用いられません。" _
                    & vbCrLf & _
                    "軸の最小値が正の数になっている事を確認してください。", vbCritical + vbQuestion, "警告"
                    
                    Dim s_inp_min_box As New ilogminvalbox
                    Call s_inp_min_box.GetInf("ｘ軸：最小値設定", "ｘ", x1_max, 1)
                    
                    x1_min = s_inp_min_box.UserInput
                    x2_min = x1_min
                    x1_max = tofliv.setLogMaxValue(1, s_x_log_base)
                    x2_max = x1_max
                End If
            
                With ActiveChart
                    .SetElement msoElementPrimaryCategoryAxisLogScale
                    .SetElement msoElementSecondaryCategoryAxisLogScale
                    .Axes(xlValue).ScaleType = xlLinear
                    
                    With .Axes(xlCategory, 1)
                        .MaximumScale = x1_max
                        .MinimumScale = x1_min
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabels.Font.Color = inp_axes_color
                        .HasTitle = True
                        Call tofliv.setTitf(1, x_axis_title, inp_axes_color)
                    End With
                    
                    With .Axes(xlCategory, 2)
                        .MaximumScale = x2_max
                        .MinimumScale = x2_min
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabelPosition = xlTickLabelPositionNone
                    End With
                    
                    With .Axes(xlValue, 1)
                        .MaximumScale = y1_max
                        .MinimumScale = y1_min
                        .MajorTickMark = xlInside
                        .ScaleType = xlLinear
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabels.Font.Color = inp_axes_color
                        .HasTitle = True
                        Call tofliv.setTitf(2, y_axis_title, inp_axes_color)
                    End With
                    
                    With .Axes(xlValue, 2)
                        .MaximumScale = y2_max
                        .MinimumScale = y2_min
                        .MajorTickMark = xlInside
                        .ScaleType = xlLinear
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabelPosition = xlTickLabelPositionNone
                    End With
                    
                    .Axes(xlCategory).HasMajorGridlines = False
                    .Axes(xlCategory).Crosses = xlMinimum
                    .Axes(xlValue).HasMajorGridlines = False
                    .Axes(xlValue).Crosses = xlMinimum
                        
                End With
            
            Case 3 'ｘ軸線形ｙ軸対数
            
                Dim t_inp_box As New ilogbasebox
                Call t_inp_box.GetInf("ｙ軸：基数設定", "ｙ", 10)
                Dim t_y_log_base As Double
                t_y_log_base = t_inp_box.UserInput
                
                If 0 < y1_min Then
                    y2_min = y1_min
                    y2_max = y1_max
                Else
                    MsgBox "対数軸に負の値は用いられません。" _
                    & vbCrLf & _
                    "軸の値が正の数になっている事を確認してください。", vbCritical + vbQuestion, "警告"
                    Dim t_inp_min_box As New ilogminvalbox
                    Call t_inp_min_box.GetInf("ｙ軸：最小値設定", "ｙ", y1_max, 1)
                    y1_min = t_inp_min_box.UserInput
                    y2_min = y1_min
                    y1_max = tofliv.setLogMaxValue(2, t_y_log_base)
                    y2_max = y1_max
                End If
                
                With ActiveChart
                    .SetElement msoElementPrimaryValueAxisLogScale
                    .SetElement msoElementSecondaryValueAxisLogScale
                    .SetElement msoElementPrimaryCategoryAxisShow
                    .SetElement msoElementSecondaryCategoryAxisShow
                    
                    With .Axes(xlCategory, 1)
                        .MaximumScale = x1_max
                        .MinimumScale = x1_min
                        .MajorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabels.Font.Color = inp_axes_color
                        .HasTitle = True
                        Call tofliv.setTitf(1, x_axis_title, inp_axes_color)
                    End With
                    
                    With .Axes(xlCategory, 2)
                        .MaximumScale = x2_max
                        .MinimumScale = x2_min
                        .MajorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabelPosition = xlTickLabelPositionNone
                    End With
                    
                    With .Axes(xlValue, 1)
                        .MaximumScale = y1_max
                        .MinimumScale = y1_min
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabels.Font.Color = inp_axes_color
                        .HasTitle = True
                        Call tofliv.setTitf(1, y_axis_title, inp_axes_color)
                    End With
                    
                    With .Axes(xlValue, 2)
                        .MaximumScale = y2_max
                        .MinimumScale = y2_min
                        .MajorTickMark = xlInside
                        .MinorTickMark = xlInside
                        .Format.Line.ForeColor.RGB = inp_axes_color
                        .TickLabelPosition = xlTickLabelPositionNone
                    End With
                    
                    .Axes(xlCategory).HasMajorGridlines = False
                    .Axes(xlCategory).Crosses = xlMinimum
                    .Axes(xlValue).HasMajorGridlines = False
                    .Axes(xlValue).Crosses = xlMinimum
                    
                End With
                
            
            Case 4 'ｘ軸線形ｙ軸線形
            
                MsgBox "両軸とも線形の場合は、" _
                & vbCrLf & _
                "「Ctrl+Shift+G」で線形軸用のマクロを実行してください。" _
                , vbInformation, "お知らせ"
        
        End Select
    
    Else
    
        ThisWorkbook.Sheets("ホーム").Protect PASSWORD:=1184
        MsgBox "申し訳ありませんが、このコンピューターでは使えません。" _
        & vbCrLf & _
        "ご購入の検討をお願い申し上げます。", vbCritical + vbQuestion, "警告"
    
    End If


End Sub



















