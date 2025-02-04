Attribute VB_Name = "FactorySetting"
Option Explicit

Public Function FactoryEx()

    If MsgBox("始めにネットワークを経由して認証をおこないます。" & vbCrLf & _
                "この操作は初期設定の一回のみです。" & vbLf & _
                    "以降はオフラインでご利用頂けます。" & vbLf & _
                        "ネットワーク認証を行いますか？", vbYesNo + vbInformation, "ネットワーク認証") = vbYes Then
                        
        Dim get_http_autho As String
        get_http_autho = tofliv.getHttpAutho
        
        If get_http_autho = "12345" Then
            Load UserForm2
            UserForm2.Caption = "利用規約"
            UserForm2.Show vbModeless
        ElseIf get_http_autho = "NetError" Then
            MsgBox "インターネット接続を確認してください。", vbExclamation
        ElseIf get_http_autho = "5??" Then
            MsgBox "サーバー側でのエラーが発生中です。" & vbLf & "しばらく時間を置いてからもう一度実行してください。", vbExclamation
        Else
            MsgBox "原因不明のエラーが発生しました。", vbCritical
        End If
        
    Else
        MsgBox "ネットワーク認証をされないとご利用いただけません。", vbExclamation
    End If

End Function

Private Sub ボタン1_Click()
    FactoryEx
End Sub

