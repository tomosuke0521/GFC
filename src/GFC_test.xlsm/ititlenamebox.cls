VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ititlenamebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private title_name As String
Private axes_name As String
Private def_title As String

Public Sub GetInf(ByVal inp_title_name As String, ByVal inp_axes_name As String, _
                    ByVal inp_def_title As String)

    title_name = inp_title_name
    axes_name = inp_axes_name
    def_title = inp_def_title

End Sub


Public Function UserInput()

    Dim axis_title As String
    axis_title = InputBox(axes_name + "軸の「タイトル」を入力してください。" _
    & vbCrLf & _
    "\で挟んだ文字が斜体となります。\は表示されません。" _
    & vbCrLf & _
    "キャンセルの場合は終了します。" _
    , title_name, Default:=def_title)
    If StrPtr(axis_title) = 0 Then
        MsgBox "キャンセルが押されたので終了します。"
        End
    End If
    
    UserInput = axis_title

End Function
