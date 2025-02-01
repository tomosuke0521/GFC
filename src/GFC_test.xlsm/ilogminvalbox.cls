VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ilogminvalbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private title_name As String
Private axes_name As String
Private max_val As Double
Private def_num As Double

Public Sub GetInf(ByVal inp_title_name As String, ByVal inp_axes As String, _
                    ByVal inp_max_val As Double, ByVal inp_def_num As Double)

    title_name = inp_title_name
    axes_name = inp_axes
    max_val = inp_max_val
    def_num = inp_def_num

End Sub

Public Function UserInput()
    Do
        Do
            Do
                Dim min_val As String
                min_val = InputBox("0より小さい値は対数軸に反映されません。" _
                & vbCrLf & _
                axes_name + "軸の「最小値」を「0～" + CStr(max_val) + "」の間で決めて下さい。" _
                & vbCrLf & _
                "キャンセルの場合は終了します。", title_name, Default:=def_num)
                If StrPtr(min_val) = 0 Then
                    MsgBox "キャンセルが押されたので終了します。"
                    End
                End If
                If Not (IsNumeric(min_val)) Then
                    MsgBox "数値で入力してください。", vbCritical
                ElseIf min_val <= 0 Then
                    MsgBox "0より大きな値で入力してください。", vbCritical
                ElseIf min_val >= max_val Then
                    MsgBox CStr(max_val) + "より小さな値で入力してください。", vbCritical
                End If
            Loop While Not (IsNumeric(min_val))
        Loop While min_val <= 0
    Loop While min_val >= max_val
    
    UserInput = min_val

End Function


