VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ilogbasebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private title_name As String
Private axes_name As String
Private def_num As Double

Public Sub GetInf(ByVal inp_title_name As String, ByVal inp_axes As String, ByVal inp_def_num As Double)

    title_name = inp_title_name
    axes_name = inp_axes
    def_num = inp_def_num

End Sub

Public Function UserInput()
    Do
        Do
            Dim log_base As String
            log_base = InputBox(axes_name + "軸の「基数」を決めて下さい。" _
            & vbCrLf & _
            "キャンセルの場合は終了します。", title_name, Default:=def_num)
            If StrPtr(log_base) = 0 Then
                MsgBox "キャンセルが押されたので終了します。"
                End
            End If
            If Not (IsNumeric(log_base)) Then
                MsgBox "数値で入力してください。", vbCritical
            ElseIf log_base <= 0 Then
                MsgBox "0より大きな値で入力してください。", vbCritical
            End If
        Loop While Not (IsNumeric(log_base))
    Loop While log_base <= 0
    
    UserInput = log_base

End Function

