VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   5190
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9540.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
    
    MsgBox "���p�K��ɓ��ӂ���Ȃ��Ɨ��p�ł��܂���B", vbInformation
    
End Sub

Private Sub CheckBox1_Click()
    
    If CheckBox1.Value Then
        If MAXSCBARVALUE <> 100 Then
            CheckBox1.Value = False
            MsgBox "���p�K���ǂ�ł��������B�i�X�N���[���o�[�����܂ŉ����Ă��������j", vbCritical
        End If
    End If
    
End Sub



Private Sub CommandButton2_Click()

    If CheckBox1.Value Then
        Unload UserForm2
        Load UserForm1
        UserForm1.Caption = "�����ݒ蒆"
        UserForm1.Show vbModeless
        
        Dim i As Integer
        For i = 0 To 100
            Application.Wait [Now()] + Rnd() / 10 / 86400
            UserForm1.ProgressBar1.Value = i
            UserForm1.Label1.Caption = i & "%"
            UserForm1.Repaint
        Next i
        
        ThisWorkbook.Sheets("�z�[��").Unprotect PASSWORD:=PASSWORD_NUMBER
        With ThisWorkbook.Sheets("�z�[��").Cells(1, 1)
            .Value = tofliv.GetHdSn()
            .Font.Color = vbWhite
        End With
        ThisWorkbook.Sheets("�z�[��").Protect PASSWORD:=PASSWORD_NUMBER
        
        MsgBox "�����ݒ肪�������܂����B", vbInformation
        Unload UserForm1
        
    Else
        MsgBox "�u���ӂ���v�Ƀ`�F�b�N���Ă�������", vbCritical
    End If
End Sub

Private Sub UserForm_Initialize()

    UserForm2.Caption = "���p�K��"

    ScrollBar1.Value = 1
    MAXSCBARVALUE = 0
    
    
    Label1.Caption = "" & vbLf & ""
    Label1.BackColor = RGB(240, 240, 240)
    Label4.Caption = "" & vbLf & ""
    Label4.BackColor = RGB(240, 240, 240)
    
    CheckBox1.Caption = "���ӂ���"
    
    CommandButton1.Font.Size = 7
    CommandButton2.Font.Size = 7
    CommandButton1.Caption = "�L�����Z��"
    CommandButton2.Caption = "����"
    
    With Label3
        .Caption = ""
        .BackStyle = fmBackStyleTransparent
        .SpecialEffect = fmSpecialEffectSunken
    End With
    
    Label2.Top = 20
    Label2.Height = 300
    Label2.Caption = "�@���̗��p�K��i�ȉ��w�{�K��x�Ƃ����j�́A�v���O���~���O�T�[�N��RyunensClub(�ȉ��w���x�Ƃ���)��" & _
    "�쐻����Excel�O���t���������ύX�c�[���i�ȉ��wGFC�x�Ƃ���)���w�����_�E�����[�h���ė��p����ҁi�ȉ��w�b�x�Ƃ����j��" & _
    "�K�p����闘�p�������߂���̂ł���B" & vbLf & vbLf & _
 _
    "��P���i�Ώۃ\�t�g�E�F�A�j" & vbLf & _
    "�P�@�{�_��ɂ����ċ����̑ΏۂƂȂ�\�t�g�E�F�A�i�ȉ��w�{�\�t�g�E�F�A�x�Ƃ����j�́A���̗��p�K�񂪎��߂��Ă���Excel�t�@�C���̂��Ƃł���B" & _
    "�܂��A�{�\�t�g�E�F�A�́A�����b�ɒ񋟂���X�V�ŋy�уo�[�W�����A�b�v�ł��܂܂��B" & vbLf & vbLf & _
 _
    "��Q���i�g�p�����j" & vbLf & _
    "�P�@���͍b�ɑ΂���GFC���b���Ǘ�����P��̃R���s���[�^�[���ɃC���X�g�[�����A�Ȃ����A���̂P��ɂ����Ă̂ݗ��p����" & _
    "���Ƃ���������i�ȉ��w�{�����x�Ƃ����j�B" & vbLf & _
    "�Q�@GFC�͌l���p�A���p���p�A�������p�A�ɖ�킸�Љ�ʔO��F�߂�ꂤ�闘�p�Ɏg�p�ł���B" & vbLf & _
    "�R�@�{�����Ɋւ��GFC�̎g�p���́A��Ɛ�I�ł���A���A�ċ����s�A���n�s�\�̂��̂Ƃ���B" & vbLf & vbLf & _
 _
    "��R���i�����A���j" & vbLf & _
    "�P�@GFC�̑S�Ă̒��쌠�͉��ɑ�����B" & vbLf & vbLf & _
 _
    "��S���i�֎~�����j" & vbLf & _
    "�P�@�{�\�t�g�E�F�A�ł���Excel�t�@�C�����O�҂ɏ��n�A�̔����邱�Ƃ��֎~����B" & vbLf & _
    "�Q�@���ς͍b�̎��R�ɍs���Ă悢���̂Ƃ��邪�A�p�X���[�h��ی삳�ꂽ�̈�����ς��Ă͂Ȃ�Ȃ��B" & vbLf & _
    "�R�@�b����O�҂���{�\�t�g�E�F�A�̋@�\�ɂ��ĕ����ꂽ�ۂ́A�ł��������A�����Љ�Ă��������B" & vbLf & vbLf & _
 _
    "��T���i�ۏ�j" & vbLf & _
    "�P�@�b���{�\�t�g�E�F�A�ɂ��������؂̑��Q�i�f�[�^����Ȃǁj�ɂ��ẮA���ɂ��̈�؂̐ӔC�͂Ȃ����̂Ƃ���B" & vbLf & _
    "�Q�@�{�\�t�g�E�F�A�Ɋւ��邢���Ȃ�s����������Ƃ��Ă����ɕۏ؂���`���͂Ȃ����̂Ƃ���B"


End Sub

Private Sub ScrollBar1_Change()

    ScrollBar1.Min = 0
    ScrollBar1.MAX = 100
    
    If MAXSCBARVALUE <= ScrollBar1.Value Then
        MAXSCBARVALUE = ScrollBar1.Value
    End If
    
    Label2.Top = 20 - ScrollBar1.Value
    
    Label2.Caption = "�@���̗��p�K��i�ȉ��w�{�K��x�Ƃ����j�́A�v���O���~���O�T�[�N��RyunensClub(�ȉ��w���x�Ƃ���)��" & _
    "�쐻����Excel�O���t���������ύX�c�[���i�ȉ��wGFC�x�Ƃ���)���w�����_�E�����[�h���ė��p����ҁi�ȉ��w�b�x�Ƃ����j��" & _
    "�K�p����闘�p�������߂���̂ł���B" & vbLf & vbLf & _
 _
    "��P���i�Ώۃ\�t�g�E�F�A�j" & vbLf & _
    "�P�@�{�_��ɂ����ċ����̑ΏۂƂȂ�\�t�g�E�F�A�i�ȉ��w�{�\�t�g�E�F�A�x�Ƃ����j�́A���̗��p�K�񂪎��߂��Ă���Excel�t�@�C���̂��Ƃł���B" & _
    "�܂��A�{�\�t�g�E�F�A�́A�����b�ɒ񋟂���X�V�ŋy�уo�[�W�����A�b�v�ł��܂܂��B" & vbLf & vbLf & _
 _
    "��Q���i�g�p�����j" & vbLf & _
    "�P�@���͍b�ɑ΂���GFC���b���Ǘ�����P��̃R���s���[�^�[���ɃC���X�g�[�����A�Ȃ����A���̂P��ɂ����Ă̂ݗ��p����" & _
    "���Ƃ���������i�ȉ��w�{�����x�Ƃ����j�B" & vbLf & _
    "�Q�@GFC�͌l���p�A���p���p�A�������p�A�ɖ�킸�Љ�ʔO��F�߂�ꂤ�闘�p�Ɏg�p�ł���B" & vbLf & _
    "�R�@�{�����Ɋւ��GFC�̎g�p���́A��Ɛ�I�ł���A���A�ċ����s�A���n�s�\�̂��̂Ƃ���B" & vbLf & vbLf & _
 _
    "��R���i�����A���j" & vbLf & _
    "�P�@GFC�̑S�Ă̒��쌠�͉��ɑ�����B" & vbLf & vbLf & _
 _
    "��S���i�֎~�����j" & vbLf & _
    "�P�@�{�\�t�g�E�F�A�ł���Excel�t�@�C�����O�҂ɏ��n�A�̔����邱�Ƃ��֎~����B" & vbLf & _
    "�Q�@���ς͍b�̎��R�ɍs���Ă悢���̂Ƃ��邪�A�p�X���[�h��ی삳�ꂽ�̈�����ς��Ă͂Ȃ�Ȃ��B" & vbLf & _
    "�R�@�b����O�҂���{�\�t�g�E�F�A�̋@�\�ɂ��ĕ����ꂽ�ۂ́A�ł��������A�����Љ�Ă��������B" & vbLf & vbLf & _
 _
    "��T���i�ۏ�j" & vbLf & _
    "�P�@�b���{�\�t�g�E�F�A�ɂ��������؂̑��Q�i�f�[�^����Ȃǁj�ɂ��ẮA���ɂ��̈�؂̐ӔC�͂Ȃ����̂Ƃ���B" & vbLf & _
    "�Q�@�{�\�t�g�E�F�A�Ɋւ��邢���Ȃ�s����������Ƃ��Ă����ɕۏ؂���`���͂Ȃ����̂Ƃ���B"

    
    
End Sub















