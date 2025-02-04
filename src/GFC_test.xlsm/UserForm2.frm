VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   5190
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9540.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim MAXSCBARVALUE As Long

Private Sub CommandButton1_Click()
    
    Unload UserForm2
    
    MsgBox "利用規約に同意されないと利用できません。", vbInformation
    
End Sub

Private Sub CheckBox1_Click()
    
    If CheckBox1.Value Then
        If MAXSCBARVALUE <> 100 Then
            CheckBox1.Value = False
            MsgBox "利用規約を読んでください。（スクロールバーを下まで下げてください）", vbCritical
        End If
    End If
    
End Sub



Private Sub CommandButton2_Click()

    If CheckBox1.Value Then
        Unload UserForm2
        Load UserForm1
        UserForm1.Caption = "初期設定中"
        UserForm1.Show vbModeless
        
        Dim i As Integer
        For i = 0 To 100
            Application.Wait [Now()] + Rnd() / 10 / 86400
            UserForm1.ProgressBar1.Value = i
            UserForm1.Label1.Caption = i & "%"
            UserForm1.Repaint
        Next i
        
        ThisWorkbook.Sheets("ホーム").Unprotect PASSWORD:=PASSWORD_NUMBER
        With ThisWorkbook.Sheets("ホーム").Cells(1, 1)
            .Value = tofliv.GetHdSn()
            .Font.Color = vbWhite
        End With
        ThisWorkbook.Sheets("ホーム").Protect PASSWORD:=PASSWORD_NUMBER
        
        MsgBox "初期設定が完了しました。", vbInformation
        Unload UserForm1
        
    Else
        MsgBox "「同意する」にチェックしてください", vbCritical
    End If
End Sub

Private Sub UserForm_Initialize()

    UserForm2.Caption = "利用規約"

    ScrollBar1.Value = 1
    MAXSCBARVALUE = 0
    
    
    Label1.Caption = "" & vbLf & ""
    Label1.BackColor = RGB(240, 240, 240)
    Label4.Caption = "" & vbLf & ""
    Label4.BackColor = RGB(240, 240, 240)
    
    CheckBox1.Caption = "同意する"
    
    CommandButton1.Font.Size = 7
    CommandButton2.Font.Size = 7
    CommandButton1.Caption = "キャンセル"
    CommandButton2.Caption = "次へ"
    
    With Label3
        .Caption = ""
        .BackStyle = fmBackStyleTransparent
        .SpecialEffect = fmSpecialEffectSunken
    End With
    
    Label2.Top = 20
    Label2.Height = 300
    Label2.Caption = "　この利用規約（以下『本規約』という）は、プログラミングサークルRyunensClub(以下『乙』という)が" & _
    "作製したExcelグラフ自動書式変更ツール（以下『GFC』という)を購入かつダウンロードして利用する者（以下『甲』という）に" & _
    "適用される利用条件を定めるものである。" & vbLf & vbLf & _
 _
    "第１条（対象ソフトウェア）" & vbLf & _
    "１　本契約において許諾の対象となるソフトウェア（以下『本ソフトウェア』という）は、この利用規約が収められているExcelファイルのことである。" & _
    "また、本ソフトウェアは、乙が甲に提供する更新版及びバージョンアップ版が含まれる。" & vbLf & vbLf & _
 _
    "第２条（使用許諾）" & vbLf & _
    "１　乙は甲に対してGFCを甲が管理する１台のコンピュータ端末にインストールし、なおかつ、その１台においてのみ利用する" & _
    "ことを許諾する（以下『本許諾』という）。" & vbLf & _
    "２　GFCは個人利用、商用利用、研究利用、に問わず社会通念上認められうる利用に使用できる。" & vbLf & _
    "３　本許諾に関わるGFCの使用権は、非独占的であり、かつ、再許諾不可、譲渡不能のものとする。" & vbLf & vbLf & _
 _
    "第３条（権利帰属）" & vbLf & _
    "１　GFCの全ての著作権は乙に属する。" & vbLf & vbLf & _
 _
    "第４条（禁止事項）" & vbLf & _
    "１　本ソフトウェアであるExcelファイルを第三者に譲渡、販売することを禁止する。" & vbLf & _
    "２　改変は甲の自由に行ってよいものとするが、パスワードや保護された領域を改変してはならない。" & vbLf & _
    "３　甲が第三者から本ソフトウェアの機能について聞かれた際は、できうる限り、乙を紹介してください。" & vbLf & vbLf & _
 _
    "第５条（保障）" & vbLf & _
    "１　甲が本ソフトウェアにより被った一切の損害（データ損壊など）については、乙にその一切の責任はないものとする。" & vbLf & _
    "２　本ソフトウェアに関するいかなる不具合が生じたとしても乙に保証する義務はないものとする。"


End Sub

Private Sub ScrollBar1_Change()

    ScrollBar1.Min = 0
    ScrollBar1.MAX = 100
    
    If MAXSCBARVALUE <= ScrollBar1.Value Then
        MAXSCBARVALUE = ScrollBar1.Value
    End If
    
    Label2.Top = 20 - ScrollBar1.Value
    
    Label2.Caption = "　この利用規約（以下『本規約』という）は、プログラミングサークルRyunensClub(以下『乙』という)が" & _
    "作製したExcelグラフ自動書式変更ツール（以下『GFC』という)を購入かつダウンロードして利用する者（以下『甲』という）に" & _
    "適用される利用条件を定めるものである。" & vbLf & vbLf & _
 _
    "第１条（対象ソフトウェア）" & vbLf & _
    "１　本契約において許諾の対象となるソフトウェア（以下『本ソフトウェア』という）は、この利用規約が収められているExcelファイルのことである。" & _
    "また、本ソフトウェアは、乙が甲に提供する更新版及びバージョンアップ版が含まれる。" & vbLf & vbLf & _
 _
    "第２条（使用許諾）" & vbLf & _
    "１　乙は甲に対してGFCを甲が管理する１台のコンピュータ端末にインストールし、なおかつ、その１台においてのみ利用する" & _
    "ことを許諾する（以下『本許諾』という）。" & vbLf & _
    "２　GFCは個人利用、商用利用、研究利用、に問わず社会通念上認められうる利用に使用できる。" & vbLf & _
    "３　本許諾に関わるGFCの使用権は、非独占的であり、かつ、再許諾不可、譲渡不能のものとする。" & vbLf & vbLf & _
 _
    "第３条（権利帰属）" & vbLf & _
    "１　GFCの全ての著作権は乙に属する。" & vbLf & vbLf & _
 _
    "第４条（禁止事項）" & vbLf & _
    "１　本ソフトウェアであるExcelファイルを第三者に譲渡、販売することを禁止する。" & vbLf & _
    "２　改変は甲の自由に行ってよいものとするが、パスワードや保護された領域を改変してはならない。" & vbLf & _
    "３　甲が第三者から本ソフトウェアの機能について聞かれた際は、できうる限り、乙を紹介してください。" & vbLf & vbLf & _
 _
    "第５条（保障）" & vbLf & _
    "１　甲が本ソフトウェアにより被った一切の損害（データ損壊など）については、乙にその一切の責任はないものとする。" & vbLf & _
    "２　本ソフトウェアに関するいかなる不具合が生じたとしても乙に保証する義務はないものとする。"

    
    
End Sub















