VERSION 5.00
Begin VB.MDIForm MDIfrm 
   BackColor       =   &H8000000C&
   Caption         =   "Y�����"
   ClientHeight    =   6645
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13200
   Icon            =   "MDIfrm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Menu File 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu Newform 
         Caption         =   "�µĴ���"
         Shortcut        =   ^N
      End
      Begin VB.Menu Exit 
         Caption         =   "�˳�"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "����(&S)"
      Begin VB.Menu InternetOptions 
         Caption         =   "Internet ѡ��"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Help 
      Caption         =   "����(&H)"
      Begin VB.Menu GitHub 
         Caption         =   "GitHub ��Դ��ַ"
         Shortcut        =   {F2}
      End
      Begin VB.Menu About 
         Caption         =   "����"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("�˲�����رձ��Ự�򿪵����б�ǩҳ�������Ҫ������", vbYesNo + vbExclamation) = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub MDIForm_Load()
'�л���IE11�ں�
CreateObject("wscript.shell").regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", "11000", "REG_DWORD"
'��������в��������������������в���ָ��������,���û���������ҳ
If Command = "" Then
frmBrowser.brwWebBrowser.GoHome
Else
frmBrowser.brwWebBrowser.Navigate Command
End If
End Sub

Private Sub Newform_Click()
Dim frmBrowser As New frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.GoHome
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub InternetOptions_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl"
End Sub

Private Sub GitHub_Click()
Dim frmBrowser As New frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate "https://github.com/VB-Studio/Y-Browser"
End Sub

Private Sub About_Click()
frmAbout.Show vbModal, Me
End Sub
