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
            Dim set_width As String
            set_width = InputBox(axes_name + "軸の「目盛の数」を選んでください。" _
            & vbCrLf & _
            "キャンセルの場合は終了します。", title_name, Default:=def_num)
            
            If StrPtr(set_width) = 0 Then
                MsgBox "キャンセルが押されたので終了します。"
                End
            End If
            
            If Not (IsNumeric(set_width)) Then
                MsgBox "数値を入力してください。", vbCritical
            ElseIf set_width <= 0 Then
                MsgBox "0より大きな値で入力してください。", vbCritical
            End If
        Loop While Not (IsNumeric(set_width))
    Loop While set_width <= 0
    
    UserInput = set_width

End Function

