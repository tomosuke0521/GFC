Attribute VB_Name = "Reset"
Option Explicit

Sub Reset()

    If MsgBox("�S�Ă̐ݒ�l���f�t�H���g�ɖ߂��܂����H", vbYesNo + vbQuestion) = vbYes Then
    
    ThisWorkbook.Sheets("�I�v�V����").Unprotect PASSWORD:=PASSWORD_NUMBER
        With ThisWorkbook.Sheets("�I�v�V����")
            Range("H9").Value = 2
            Range("H10").Value = 2
            Range("H13").Value = 2
            Range("H14").Value = 2
            Range("H17").Value = 1
            Range("H22").Value = 1
            Range("H27").Value = 1
        End With
        Call setSampleColor
    ThisWorkbook.Sheets("�I�v�V����").Protect PASSWORD:=PASSWORD_NUMBER
    
    Else
    
        Exit Sub
    
    End If
    

End Sub
