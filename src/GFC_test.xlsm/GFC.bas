Attribute VB_Name = "GFC"
Option Explicit


Sub GraphSeriesAuto()
Attribute GraphSeriesAuto.VB_ProcData.VB_Invoke_Func = "G\n14"

    
    'Application.ScreenUpdating = False
    
    With ThisWorkbook.Sheets("�I�v�V����")
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
    HD_S_number = ThisWorkbook.Sheets("�z�[��").Cells(1, 1).Value
    
    
    'Application.ScreenUpdating = True
    

    If tofliv.GetHdSn() = HD_S_number Then
    
        '���擾�Z�N�V����
        On Error Resume Next
            Dim series_count As Long
            series_count = ActiveChart.SeriesCollection.Count
            If Err.Number <> 0 Then
                MsgBox "�ύX����O���t��I�����Ă��������B", vbOKOnly + vbCritical, "�ǂݍ��݃G���["
                Exit Sub
            End If
        On Error GoTo 0
        
        Dim srs_flag As Boolean
            srs_flag = True
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
                srs_flag = False
                If Err.Number <> 0 Then
                    MsgBox "�Ή����Ă���̂́A�u�U�z�}�v�݂̂ł�", vbCritical
                    Exit Sub
                End If
            On Error GoTo 0
        End If
        
        Dim x_max As Double, x_min As Double
        Dim y_max As Double, y_min As Double
        Dim gatv As getautotickvalue
        set gatv = new getautotickvalue
        call gatv.GetAutoMaxMinTickValue(srs_flag)
        x_max = gatv.xAoutMaxTickValue
        x_min = gatv.xAoutminTickValue
        y_max = gatv.yAoutMaxTickValue
        y_min = gatv.yAoutMinTickValue
        Set gatv = Nothing

        
        '���̃^�C�g���̎擾
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
        Call x_width_box.GetInf("�����F�ڐ��萔�I��", "��", get_x_width)
        Dim inp_x_width As String
        inp_x_width = x_width_box.UserInput
        Dim x_width As Double
        x_width = (x_max - x_min) / (inp_x_width + 1)
        
        Dim y_width_box As New ilintickvalbox
        Call y_width_box.GetInf("�����F�ڐ��萔�I��", "��", get_y_width)
        Dim inp_y_width As Double
        inp_y_width = y_width_box.UserInput
        Dim y_width As Double
        y_width = (y_max - y_min) / (inp_y_width + 1)
        
        Dim c_x_axis_title As New ititlenamebox
        Call c_x_axis_title.GetInf("�����F�^�C�g���ݒ�", "��", get_x_axis_title)
        Dim x_axis_title As String
        x_axis_title = c_x_axis_title.UserInput
        
        Dim c_y_axis_title As New ititlenamebox
        Call c_y_axis_title.GetInf("�����F�^�C�g���ݒ�", "��", get_y_axis_title)
        Dim y_axis_title As String
        y_axis_title = c_y_axis_title.UserInput
        
        
        
'------------�����ݒ�Z�N�V����------------

        Dim inp_axes_color As Long
        inp_axes_color = RGB(rgb_r, rgb_g, rgb_b)
        
        Call tofliv.Mitm(x_mitm_flag, y_mitm_flag, x_width, y_width)
        
        With ActiveChart
        
            '��P����xlCategory
            With .Axes(xlCategory, 1)
                .MajorTickMark = xlInside
                .Format.Line.ForeColor.RGB = inp_axes_color
                On Error Resume Next
                    .MaximumScale = x_max
                    .MinimumScale = x_min
                    
                    .MajorUnit = x_width
                     If Err.Number <> 0 Then
                        MsgBox "�\���󂠂�܂��񂪁A�ΐ����ɂ͑Ή����Ă���܂���B", vbOKOnly + vbCritical, "�ǂݍ��݃G���["
                        Exit Sub
                     End If
                On Error GoTo 0
                    
                Dim xtick_numfor As String
                xtick_numfor = ActiveChart.Axes(xlCategory, 1).TickLabels.NumberFormatLocal
                Call tofliv.Tnf(1, xtick_numfor, xdp_g, x_width)
                
                Call tofliv.setTitf(1, x_axis_title, inp_axes_color)
                
                .TickLabels.Font.Color = inp_axes_color
            End With
            
            '��P����
            With .Axes(xlValue, 1)
                .Format.Line.ForeColor.RGB = inp_axes_color
                On Error Resume Next
                    .MaximumScale = y_max
                    .MinimumScale = y_min
                    
                    .MajorUnit = y_width
                    If Err.Number <> 0 Then
                        MsgBox "�\���󂠂�܂��񂪁A�ΐ����ɂ͑Ή����Ă���܂���B", vbOKOnly + vbCritical, "�ǂݍ��݃G���["
                        Exit Sub
                    End If
                On Error GoTo 0
                
                Dim ytick_numfor As String
                ytick_numfor = ActiveChart.Axes(xlValue, 1).TickLabels.NumberFormatLocal
                Call tofliv.Tnf(2, ytick_numfor, ydp_g, y_width)
                
                Call tofliv.setTitf(2, y_axis_title, inp_axes_color)
                
                .TickLabels.Font.Color = inp_axes_color
            End With
            
            '��Q����
            With .Axes(xlCategory, 2)
                .TickLabelPosition = xlTickLabelPositionNone
                .MajorTickMark = xlInside
                On Error Resume Next
                    .MaximumScale = x_max
                    .MinimumScale = x_min
                    .MajorUnit = x_width
                    If Err.Number <> 0 Then
                        MsgBox "�\���󂠂�܂��񂪁A�ΐ����ɂ͑Ή����Ă���܂���B", vbOKOnly + vbCritical, "�ǂݍ��݃G���["
                        Exit Sub
                    End If
                On Error GoTo 0
                
                .Format.Line.ForeColor.RGB = inp_axes_color
            End With
            
            '��Q����
            With .Axes(xlValue, 2)
                .TickLabelPosition = xlTickLabelPositionNone
                .TickLabelPosition = xlTickLabelPositionNone
                .MajorTickMark = xlInside
                On Error Resume Next
                    .MaximumScale = y_max
                    .MinimumScale = y_min
                    .MajorUnit = y_width
                     If Err.Number <> 0 Then
                        MsgBox "�\���󂠂�܂��񂪁A�ΐ����ɂ͑Ή����Ă���܂���B", vbOKOnly + vbCritical, "�ǂݍ��݃G���["
                        Exit Sub
                    End If
                On Error GoTo 0
                
                .Format.Line.ForeColor.RGB = inp_axes_color
            End With
            
            '�O���t�S�̂̏����ݒ�
        
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
            
            If 1 = ThisWorkbook.Sheets("�I�v�V����").Range("M12").Value Then
                Call tofliv.setlegend
            End If
            
            If 1 = ThisWorkbook.Sheets("�I�v�V����").Range("M16").Value Then
                Call tofliv.setSeriesMarkerFormat
            End If
            
        End With
    
    Else
    
        MsgBox "�\���󂠂�܂��񂪁A���̃R���s���[�^�[�ł͎g���܂���B" _
        & vbCrLf & _
        "���w���̌��������肢�\���グ�܂��B", vbCritical + vbQuestion, "�x��"
    
    End If
    
    
End Sub





'-----------------------------------���O�o�[�W����-----------------------------------



Sub GraphSeriesLogAuto()
Attribute GraphSeriesLogAuto.VB_ProcData.VB_Invoke_Func = "A\n14"


    With ThisWorkbook.Sheets("�I�v�V����")
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
    
    If tofliv.GetHdSn() = ThisWorkbook.Sheets("�z�[��").Cells(1, 1).Value Then
        
         '���擾�Z�N�V����
        On Error Resume Next
            Dim series_count As Long
            series_count = ActiveChart.SeriesCollection.Count
            If Err.Number <> 0 Then
                MsgBox "�ύX����O���t��I�����Ă��������B", vbOKOnly + vbCritical, "�ǂݍ��݃G���["
                Exit Sub
            End If
        On Error GoTo 0
        
        
        '�n��̌��̔���ƌn��̒ǉ�
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
        
        
        '���̍ŏ��l�ƍő�l�̎擾
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
        Call c_x_axis_title.GetInf("�����F�^�C�g���ݒ�", "��", get_x_axis_title)
        Dim x_axis_title As String
        x_axis_title = c_x_axis_title.UserInput
        Dim c_y_axis_title As New ititlenamebox
        Call c_y_axis_title.GetInf("�����F�^�C�g���ݒ�", "��", get_y_axis_title)
        Dim y_axis_title As String
        y_axis_title = c_y_axis_title.UserInput
        
        
        Dim sele_log_axis As Long
        
        If MsgBox("������ΐ��\���ɂ��܂����H" & vbCrLf & "�u�������v�̏ꍇ�ɂ͐��`���ɂ��܂��B", vbYesNo, "�����I��") = vbYes Then
        
            If MsgBox("������ΐ��\���ɂ��܂����H" & vbCrLf & "�u�������v�̏ꍇ�ɂ͐��`���ɂ��܂��B", vbYesNo, "�����I��") = vbYes Then
                
                sele_log_axis = 1
            
            Else
                
                sele_log_axis = 2
            
            End If
                
        Else
            If MsgBox("������ΐ��\���ɂ��܂����H" & vbCrLf & "�u�������v�̏ꍇ�ɂ͐��`���ɂ��܂��B", vbYesNo, "�����I��") = vbYes Then
                
                sele_log_axis = 3
            Else
            
                sele_log_axis = 4
                
            End If
        End If
        
        Select Case sele_log_axis
            
            Case 1 '�����ΐ������ΐ�
            
                Dim p_x_inp_box As New ilogbasebox
                Call p_x_inp_box.GetInf("�����F��ݒ�", "x", 10)
                Dim p_x_log_base As Double
                p_x_log_base = p_x_inp_box.UserInput
                
                Dim p_y_inp_box As New ilogbasebox
                Call p_y_inp_box.GetInf("�����F��ݒ�", "��", 10)
                Dim p_y_log_base As Double
                p_y_log_base = p_y_inp_box.UserInput
                
                x1_min = gtv.xLogMinTickValue(p_x_log_base)
                    
                If 0 < x1_min Then
                    x2_min = x1_min
                    x2_max = x1_max
                Else
                    MsgBox "�ΐ�����0�ȉ��̒l�͗p�����܂���B" _
                    & vbCrLf & _
                    "���̍ŏ��l�����̐��ɂȂ��Ă��鎖���m�F���Ă��������B", vbCritical + vbQuestion, "�x��"
                    
                    Dim p_x_inp_min_box As New ilogminvalbox
                    Call p_x_inp_min_box.GetInf("�����F�ŏ��l�ݒ�", "��", x1_max, 1)
                    
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
                    MsgBox "�ΐ����ɕ��̒l�͗p�����܂���B" _
                    & vbCrLf & _
                    "���̒l�����̐��ɂȂ��Ă��鎖���m�F���Ă��������B", vbCritical + vbQuestion, "�x��"
                    Dim p_y_inp_min_box As New ilogminvalbox
                    Call p_y_inp_min_box.GetInf("�����F�ŏ��l�ݒ�", "��", y1_max, 1)
                    y1_min = p_y_inp_min_box.UserInput
                    y2_min = y1_min
                    y1_max = tofliv.setLogMaxValue(2, p_y_log_base)
                    y2_max = y1_max
                End If
                
                
                '�O���t�����ݒ�
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
                
                
                
            
            Case 2 '�����ΐ��������`
            
                Dim s_inp_box As New ilogbasebox
                Call s_inp_box.GetInf("�����F��ݒ�", "x", 10)
                Dim s_x_log_base As Double
                s_x_log_base = s_inp_box.UserInput
                    
                If 0 < x1_min Then
                    x2_min = x1_min
                    x2_max = x1_max
                Else
                    MsgBox "�ΐ�����0�ȉ��̒l�͗p�����܂���B" _
                    & vbCrLf & _
                    "���̍ŏ��l�����̐��ɂȂ��Ă��鎖���m�F���Ă��������B", vbCritical + vbQuestion, "�x��"
                    
                    Dim s_inp_min_box As New ilogminvalbox
                    Call s_inp_min_box.GetInf("�����F�ŏ��l�ݒ�", "��", x1_max, 1)
                    
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
            
            Case 3 '�������`�����ΐ�
            
                Dim t_inp_box As New ilogbasebox
                Call t_inp_box.GetInf("�����F��ݒ�", "��", 10)
                Dim t_y_log_base As Double
                t_y_log_base = t_inp_box.UserInput
                
                If 0 < y1_min Then
                    y2_min = y1_min
                    y2_max = y1_max
                Else
                    MsgBox "�ΐ����ɕ��̒l�͗p�����܂���B" _
                    & vbCrLf & _
                    "���̒l�����̐��ɂȂ��Ă��鎖���m�F���Ă��������B", vbCritical + vbQuestion, "�x��"
                    Dim t_inp_min_box As New ilogminvalbox
                    Call t_inp_min_box.GetInf("�����F�ŏ��l�ݒ�", "��", y1_max, 1)
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
                
            
            Case 4 '�������`�������`
            
                MsgBox "�����Ƃ����`�̏ꍇ�́A" _
                & vbCrLf & _
                "�uCtrl+Shift+G�v�Ő��`���p�̃}�N�������s���Ă��������B" _
                , vbInformation, "���m�点"
        
        End Select
    
    Else
    
        ThisWorkbook.Sheets("�z�[��").Protect PASSWORD:=1184
        MsgBox "�\���󂠂�܂��񂪁A���̃R���s���[�^�[�ł͎g���܂���B" _
        & vbCrLf & _
        "���w���̌��������肢�\���グ�܂��B", vbCritical + vbQuestion, "�x��"
    
    End If


End Sub



















