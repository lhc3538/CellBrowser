VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 

Private Sub Command1_Click()
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.gif" & String$(250, 0) '����Ĭ��Ҫ���ļ�����չ��
    file.nMaxFile = 255 '��ʾ�ļ����ĳ���
    file.lpstrFileTitle = String$(255, 0) '�򿪶Ի���ı���
    file.nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
    file.lpstrInitialDir = Environ$("WinDir") '�����̷�
    file.lpstrFilter = "ͼƬ�ļ�" & Chr$(0) & "*.gif;& Chr$(0) & Chr$(0)" '�򿪵��ļ�����
    file.nFilterIndex = 1
    file.lpstrTitle = "���ļ�"
    lResult = GetOpenFileName(file) 'ȡ���ļ���
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        MsgBox "�򿪵��ļ���Ϊ " & sFileTitle & Chr$(13) & Chr$(10) & "·��Ϊ: " & sFile, , "Open"
    End If

End Sub
