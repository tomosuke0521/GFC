Attribute VB_Name = "FactorySetting"
Option Explicit

Public Function FactoryEx()

    If MsgBox("�n�߂Ƀl�b�g���[�N���o�R���ĔF�؂������Ȃ��܂��B" & vbCrLf & _
                "���̑���͏����ݒ�̈��݂̂ł��B" & vbLf & _
                    "�ȍ~�̓I�t���C���ł����p�����܂��B" & vbLf & _
                        "�l�b�g���[�N�F�؂��s���܂����H", vbYesNo + vbInformation, "�l�b�g���[�N�F��") = vbYes Then
                        
        Dim get_http_autho As String
        get_http_autho = tofliv.getHttpAutho
        
        If get_http_autho = "12345" Then
            Load UserForm2
            UserForm2.Caption = "���p�K��"
            UserForm2.Show vbModeless
        ElseIf get_http_autho = "NetError" Then
            MsgBox "�C���^�[�l�b�g�ڑ����m�F���Ă��������B", vbExclamation
        ElseIf get_http_autho = "5??" Then
            MsgBox "�T�[�o�[���ł̃G���[���������ł��B" & vbLf & "���΂炭���Ԃ�u���Ă��������x���s���Ă��������B", vbExclamation
        Else
            MsgBox "�����s���̃G���[���������܂����B", vbCritical
        End If
        
    Else
        MsgBox "�l�b�g���[�N�F�؂�����Ȃ��Ƃ����p���������܂���B", vbExclamation
    End If

End Function

Private Sub �{�^��1_Click()
    FactoryEx
End Sub

