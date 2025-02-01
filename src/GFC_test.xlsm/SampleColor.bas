Attribute VB_Name = "SampleColor"
Option Explicit

Sub setSampleColor()

    ThisWorkbook.Sheets("�I�v�V����").Unprotect PASSWORD:=PASSWORD_NUMBER
    
    With ThisWorkbook.Sheets("�I�v�V����")
        
        Dim rgb_r, rgb_g, rgb_b As Integer
        rgb_r = .Range("H17").Value - 1
        rgb_g = .Range("H22").Value - 1
        rgb_b = .Range("H27").Value - 1
        
        .Range("G32").Interior.Color = RGB(rgb_r, rgb_g, rgb_b)
        
    End With
    
    ThisWorkbook.Sheets("�I�v�V����").Protect PASSWORD:=PASSWORD_NUMBER
    


End Sub
