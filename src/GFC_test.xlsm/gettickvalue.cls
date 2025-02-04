VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gettickvalue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private min_t_x As Double
Private max_t_x As Double
Private min_t_y As Double
Private max_t_y As Double

Private log_min_t_x As Double
Private log_max_t_x As Double
Private log_min_t_y As Double
Private log_max_t_y As Double

Dim nouse As Variant

Private Const THRESHOLD_SHOW_PROBAR As Long = 10000



Public Function Name()
    
End Function


Private Sub Class_Initialize()

    Application.ScreenUpdating = False
    
    On Error Resume Next
        Dim series_count As Long
        series_count = ActiveChart.SeriesCollection.Count
        If Err.Number <> 0 Then
            MsgBox "変更するグラフを選択してください。", vbOKOnly + vbCritical, "読み込みエラー"
            Exit Sub
        End If
    On Error GoTo 0
    If ActiveChart.SeriesCollection(2).Name = "Non" Then
        series_count = 1
    End If

    Dim srcc As Integer
    For srcc = 1 To series_count
    
        
        Dim i As Long
        
        'x軸最大値最小値取得
        Dim myarrx() As Variant
        myarrx = ActiveChart.SeriesCollection(srcc).XValues
        
        Dim ele_count As Long
        ele_count = UBound(myarrx)

        Dim max_arr_x(1 To 100) As Double
        max_arr_x(srcc) = Application.WorksheetFunction.Max(myarrx)
        Dim min_arr_x(1 To 100) As Double
        min_arr_x(srcc) = Application.WorksheetFunction.Min(myarrx)
        
        'y軸最大値最小値取得
        Dim myarry As Variant
        myarry = ActiveChart.SeriesCollection(srcc).Values
        
        Dim max_arr_y(1 To 100) As Double
        max_arr_y(srcc) = Application.WorksheetFunction.Max(myarry)
        Dim min_arr_y(1 To 100) As Double
        min_arr_y(srcc) = Application.WorksheetFunction.Min(myarry)
        Dim ele_arr(1 To 100) As Double
        ele_arr(srcc) = ele_count
        
    Next srcc
    
    
    With Application.WorksheetFunction
        Dim max_x As Double, min_x As Double
        max_x = .Max(max_arr_x)
        min_x = .Min(min_arr_x)

        Dim max_y As Double, min_y As Double
        max_y = .Max(max_arr_y)
        min_y = .Min(min_arr_y)

        Dim max_ele As Double, min_ele As Double
        max_ele = .Max(ele_arr)
        min_ele = .Min(ele_arr)
    End With
   
    
    Dim mar_x As Double: mar_x = 0.05
    Dim mar_y As Double: mar_y = 0.1
    
    min_t_x = min_x - (max_x - min_x) * mar_x
    max_t_x = max_x + (max_x - min_x) * mar_x
    
    min_t_y = min_y - (max_y - min_y) * mar_y
    max_t_y = max_y + (max_y - min_y) * mar_y
    
    log_min_t_x = min_x
    log_max_t_x = max_x
    log_min_t_y = min_y
    log_max_t_y = max_y
    
    Application.ScreenUpdating = True

End Sub

Public Function xMinTickValue()
    xMinTickValue = min_t_x
End Function

Public Function xMaxTickValue()
    xMaxTickValue = max_t_x
End Function

Public Function yMinTickValue()
    yMinTickValue = min_t_y
End Function

Public Function yMaxTickValue()
    yMaxTickValue = max_t_y
End Function

Public Function xLogMinTickValue(logbase As Double)
    xLogMinTickValue = log_min_t_x / logbase
End Function

Public Function xLogMaxTickValue(logbase As Double)
    xLogMaxTickValue = log_max_t_x * logbase
End Function

Public Function yLogMinTickValue(logbase As Double)
    yLogMinTickValue = log_min_t_y / logbase
End Function

Public Function yLogMaxTickValue(logbase As Double)
    yLogMaxTickValue = log_max_t_y * logbase
End Function




