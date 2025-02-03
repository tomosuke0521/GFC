VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ilintickvalbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

private m_x1_max As Double, m_x1_min As Double
Private m_x2_max As Double, m_x2_min As Double
private m_y1_max As Double, m_y1_min As Double
Private m_y2_max As Double, m_y2_min As Double
private m_x_max As Double, m_x_min As Double
private m_y_max As Double, m_y_min As Double
        
Private Sub GetAutoMaxMinTickValue(inp_Is_OneSeries As Boolean)
        
        '軸の境界値が自動かどうかの判定
        Dim x1_flag As Boolean
        x1_flag = ActiveChart.Axes(xlCategory, 1).MinimumScaleIsAuto
        Dim y1_flag As Boolean
        y1_flag = ActiveChart.Axes(xlValue, 1).MinimumScaleIsAuto
        Dim x2_flag As Boolean
        x2_flag = ActiveChart.Axes(xlCategory, 2).MinimumScaleIsAuto
        Dim y2_flag As Boolean
        y2_flag = ActiveChart.Axes(xlValue, 2).MinimumScaleIsAuto
        
        m_x1_min = ActiveChart.Axes(xlCategory, 1).MinimumScale
        m_y1_max = ActiveChart.Axes(xlValue, 1).MaximumScale
        m_y1_min = ActiveChart.Axes(xlValue, 1).MinimumScale
        m_x2_max = ActiveChart.Axes(xlCategory, 2).MaximumScale
        m_x2_min = ActiveChart.Axes(xlCategory, 2).MinimumScale
        m_y2_max = ActiveChart.Axes(xlValue, 2).MaximumScale
        m_y2_min = ActiveChart.Axes(xlValue, 2).MinimumScale
        
        If 1 = ThisWorkbook.Sheets("オプション").Range("M8").Value Or (x1_flag And y1_flag) Then
            Dim gtv As gettickvalue
            Set gtv = New gettickvalue
            m_x_min = gtv.xMinTickValue
            m_x_max = gtv.xMaxTickValue
            m_y_min = gtv.yMinTickValue
            m_y_max = gtv.yMaxTickValue
            Set gtv = Nothing
        Else
            If x1_flag And inp_Is_OneSeries Then
                If m_x1_max < m_x2_max Then
                    m_x_max = m_x2_max
                Else
                    m_x_max = m_x1_max
                End If
                If m_x1_min < m_x2_min Then
                    m_x_min = m_x1_min
                Else
                    m_x_min = m_x2_min
                End If
            End If
            
            If y1_flag And inp_Is_OneSeries Then
                If m_y1_max < m_y2_max Then
                    m_y_max = m_y2_max
                Else
                    m_y_max = m_y1_max
                End If
                If m_y1_min < m_y2_min Then
                    m_y_min = m_y1_min
                Else
                    m_y_min = m_y2_min
                End If
            Else
                If m_x1_max <> m_x2_max Or m_x1_min <> m_x2_min And inp_Is_OneSeries Then
                    If MsgBox("第１軸に合わせますか？" _
                    & vbCrLf & _
                    "「いいえ」の場合には第２軸に合わせます。", vbYesNo, "x軸選択") = vbYes Then
                        m_x_max = m_x1_max
                        m_x_min = m_x1_min
                        m_y_max = m_y1_max
                        m_y_min = m_y1_min
                    Else
                        m_x_max = m_x2_max
                        m_x_min = m_x2_min
                        m_y_max = m_y2_max
                        m_y_min = m_y2_min
                    End If
                Else
                    m_x_max = m_x1_max
                    m_x_min = m_x1_min
                End If
            
                If m_y1_max <> m_y2_max Or m_y1_min <> m_y2_min And inp_Is_OneSeries Then
                    If MsgBox("第１軸に合わせますか？" _
                    & vbCrLf & _
                    "「いいえ」の場合には第２軸に合わせます", vbYesNo, "y軸選択") = vbYes Then
                        m_y_max = m_y1_max
                        m_y_min = m_y1_min
                    Else
                        m_y_max = m_y2_max
                        m_y_min = m_y2_min
                    End If
                Else
                    m_y_max = m_y1_max
                    m_y_min = m_y1_min
                End If
            End If
        End If

end sub

private function xAoutMaxTickValue()
    xAoutMaxTickValue = m_x_max
end function

private function xAoutMinTickValue()
    xAoutMinTickValue = m_x_min
end function

private function yAoutMaxTickValue()
    yAoutMaxTickValue = m_y_max
end function

private function yAoutMinTickValue()
    yAoutMinTickValue = m_y_min
end function