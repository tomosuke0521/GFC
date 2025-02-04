Attribute VB_Name = "SampleColor"
Option Explicit

Sub setSampleColor()

    ThisWorkbook.Sheets("オプション").Unprotect PASSWORD:=PASSWORD_NUMBER
    
    With ThisWorkbook.Sheets("オプション")
        
        Dim rgb_r, rgb_g, rgb_b As Integer
        rgb_r = .Range("H17").Value - 1
        rgb_g = .Range("H22").Value - 1
        rgb_b = .Range("H27").Value - 1
        
        .Range("G32").Interior.Color = RGB(rgb_r, rgb_g, rgb_b)
        
    End With
    
    ThisWorkbook.Sheets("オプション").Protect PASSWORD:=PASSWORD_NUMBER
    


End Sub
