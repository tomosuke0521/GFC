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
    
        Dim myarrx() As Variant
        myarrx = ActiveChart.SeriesCollection(srcc).XValues
        Dim flag As Boolean
        flag = True
        Dim ele_count As Long
        ele_count = UBound(myarrx) + 1
        Dim arrx(2000000) As String
        ele_count = 1
        Call tofliv.showProBar("ｘ軸値読み込み中", THRESHOLD_SHOW_PROBAR, UBound(myarrx))
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
            
            If ele_count Mod 1000 = 0 Then
                DoEvents
                UserForm1.ProgressBar1.Value = ele_count / UBound(myarrx) * 100
                UserForm1.Label1.Caption = ele_count / UBound(myarrx) * 100 & "%"
                UserForm1.Repaint

            End If
            
        Loop While flag
        
        Unload UserForm1
        
        Dim i As Long
        
        Dim max_arr_x(1 To 100) As Double
        Dim min_arr_x(1 To 100) As Double
        
        Call tofliv.showProBar("ｘ軸値ソート中", THRESHOLD_SHOW_PROBAR, UBound(myarrx))
        
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
            
            If i Mod 1000 = 0 Then
                DoEvents
                UserForm1.ProgressBar1.Value = i / ele_count * 100
                UserForm1.Label1.Caption = i / ele_count * 100 & "%"
                UserForm1.Repaint
            End If
            
        Next i
        
        Unload UserForm1
        
        
        Call tofliv.showProBar("y軸値取り込み中", THRESHOLD_SHOW_PROBAR, UBound(myarrx))
        
        Dim myarry As Variant
        myarry = ActiveChart.SeriesCollection(srcc).Values
        
        Dim arr_y(2000000) As String
        For i = 1 To ele_count
            arr_y(i) = myarry(i)
            If i Mod 1000 = 0 Then
                DoEvents
                UserForm1.ProgressBar1.Value = i / ele_count * 100
                UserForm1.Label1.Caption = i / ele_count * 100 & "%"
                UserForm1.Repaint
            End If
        Next i
        
        Unload UserForm1
        
        Dim max_arr_y(1 To 100) As Double
        Dim min_arr_y(1 To 100) As Double
        Dim ele_arr(1 To 100) As Double
        Call tofliv.showProBar("y軸値ソート中", THRESHOLD_SHOW_PROBAR, UBound(myarrx))
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
            If i Mod 1000 = 0 Then
                DoEvents
                UserForm1.ProgressBar1.Value = i / ele_count * 100
                UserForm1.Label1.Caption = i / ele_count * 100 & "%"
                UserForm1.Repaint
            End If
        Next i
        
        Unload UserForm1
        
        ele_arr(srcc) = ele_count
        
    Next srcc
    
    
    
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




