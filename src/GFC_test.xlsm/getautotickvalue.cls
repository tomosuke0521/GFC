VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "getautotickvalue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Is_x1Auto As Boolean
Private m_Is_x2Auto As Boolean
Private m_Is_y1Auto As Boolean
Private m_Is_y2Auto As Boolean

Private m_x_max As Double
Private m_x_min As Double
Private m_y_max As Double
Private m_y_min As Double

Private non As Variant


Public Sub IsAutoTickValue()
    Dim Is_x1MaxTickValueAuto As Boolean
    Is_x1MaxTickValueAuto = ActiveChart.Axes(xlCategory, 1).MaximumScaleIsAuto
    Dim Is_x1MinTickValueAuto As Boolean
    Is_x1MinTickValueAuto = ActiveChart.Axes(xlCategory, 1).MinimumScaleIsAuto

    If Is_x1MaxTickValueAuto Or Is_x1MinTickValueAuto Then
        m_Is_x1Auto = True
    Else
        m_Is_x1Auto = False
    End If

    Dim Is_x2MaxTickValueAuto As Boolean
    Is_x2MaxTickValueAuto = ActiveChart.Axes(xlCategory, 2).MaximumScaleIsAuto
    Dim Is_x2MinTickValueAuto As Boolean
    Is_x2MinTickValueAuto = ActiveChart.Axes(xlCategory, 2).MinimumScaleIsAuto
    
    If Is_x2MaxTickValueAuto Or Is_x2MinTickValueAuto Then
        m_Is_x2Auto = True
    Else
        m_Is_x2Auto = False
    End If

    Dim Is_y1MaxTickValueAuto As Boolean
    Is_y1MaxTickValueAuto = ActiveChart.Axes(xlValue, 1).MaximumScaleIsAuto
    Dim Is_y1MinTickValueAuto As Boolean
    Is_y1MinTickValueAuto = ActiveChart.Axes(xlValue, 1).MinimumScaleIsAuto
    
    If Is_y1MaxTickValueAuto Or Is_y1MinTickValueAuto Then
        m_Is_y1Auto = True
    Else
        m_Is_y1Auto = False
    End If

    Dim Is_y2MaxTickValueAuto As Boolean
    Is_y2MaxTickValueAuto = ActiveChart.Axes(xlValue, 2).MaximumScaleIsAuto
    Dim Is_y2MinTickValueAuto As Boolean
    Is_y2MinTickValueAuto = ActiveChart.Axes(xlValue, 2).MinimumScaleIsAuto
    
    If Is_y2MaxTickValueAuto Or Is_y2MinTickValueAuto Then
        m_Is_y2Auto = True
    Else
        m_Is_y2Auto = False
    End If

End Sub


Public Sub GetAutoMaxMinTickValue(ByVal inp_Is_oneSeries As Boolean)
    
    '軸の境界値が自動かどうかの判定
    call IsAutoTickValue()
    '------
    
    Dim x1_max As Double, x1_min As Double
    Dim y1_max As Double, y1_min As Double
    Dim x2_max As Double, x2_min As Double
    Dim y2_max As Double, y2_min As Double
    On Error Resume Next
        '４軸の最大値と最小値取得
        x1_min = ActiveChart.Axes(xlCategory, 1).MinimumScale
        x1_max = ActiveChart.Axes(xlCategory, 1).MaximumScale
        y1_max = ActiveChart.Axes(xlValue, 1).MaximumScale
        y1_min = ActiveChart.Axes(xlValue, 1).MinimumScale
        x2_max = ActiveChart.Axes(xlCategory, 2).MaximumScale
        x2_min = ActiveChart.Axes(xlCategory, 2).MinimumScale
        y2_max = ActiveChart.Axes(xlValue, 2).MaximumScale
        y2_min = ActiveChart.Axes(xlValue, 2).MinimumScale
        If Err.Number <> 0 Then
            MsgBox "対応しているのは、「散布図」のみです", vbCritical
            Exit Sub
        End If
    On Error GoTo 0
    
    
    If 1 = ThisWorkbook.Sheets("オプション").Range("M8").Value Or (m_Is_x1Auto And m_Is_y1Auto) Then
        Dim gtv As gettickvalue
        Set gtv = New gettickvalue
        m_x_min = gtv.xMinTickValue
        m_x_max = gtv.xMaxTickValue
        m_y_min = gtv.yMinTickValue
        m_y_max = gtv.yMaxTickValue
        Set gtv = Nothing
    Else
        If m_Is_x1Auto And inp_Is_oneSeries Then
            If x1_max < x2_max Then
                m_x_max = x2_max
            Else
                m_x_max = x1_max
            End If
            If x1_min < x2_min Then
                m_x_min = x1_min
            Else
                m_x_min = x2_min
            End If
        End If
        
        If m_Is_y1Auto And inp_Is_oneSeries Then
            If y1_max < y2_max Then
                m_y_max = y2_max
            Else
                m_y_max = y1_max
            End If
            If y1_min < y2_min Then
                m_y_min = y1_min
            Else
                m_y_min = y2_min
            End If
        Else
            If x1_max <> x2_max Or x1_min <> x2_min And inp_Is_oneSeries Then
                If MsgBox("第１軸に合わせますか？" _
                & vbCrLf & _
                "「いいえ」の場合には第２軸に合わせます。", vbYesNo, "x軸選択") = vbYes Then
                    m_x_max = x1_max
                    m_x_min = x1_min
                    m_y_max = y1_max
                    m_y_min = y1_min
                Else
                    m_x_max = x2_max
                    m_x_min = x2_min
                    m_y_max = y2_max
                    m_y_min = y2_min
                End If
            Else
                m_x_max = x1_max
                m_x_min = x1_min
            End If
        
            If y1_max <> y2_max Or y1_min <> y2_min And inp_Is_oneSeries Then
                If MsgBox("第１軸に合わせますか？" _
                & vbCrLf & _
                "「いいえ」の場合には第２軸に合わせます", vbYesNo, "y軸選択") = vbYes Then
                    m_y_max = y1_max
                    m_y_min = y1_min
                Else
                    m_y_max = y2_max
                    m_y_min = y2_min
                End If
            Else
                m_y_max = y1_max
                m_y_min = y1_min
            End If
        End If
    End If
End Sub



Public Function xMaxAutoTickValue()
    xMaxAutoTickValue = m_x_max
End Function

Public Function xMinAutoTickValue()
    xMinAutoTickValue = m_x_min
End Function

Public Function yMaxAutoTickValue()
    yMaxAutoTickValue = m_y_max
End Function

Public Function yMinAutoTickValue()
    yMinAutoTickValue = m_y_min
End Function




























