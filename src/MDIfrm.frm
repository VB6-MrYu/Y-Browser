VERSION 5.00
Begin VB.MDIForm MDIfrm 
   BackColor       =   &H8000000C&
   Caption         =   "Y浏览器"
   ClientHeight    =   6645
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13200
   Icon            =   "MDIfrm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Menu File 
      Caption         =   "文件(&F)"
      Begin VB.Menu Newform 
         Caption         =   "新的窗口"
         Shortcut        =   ^N
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "设置(&S)"
      Begin VB.Menu InternetOptions 
         Caption         =   "Internet 选项"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu GitHub 
         Caption         =   "GitHub 开源地址"
         Shortcut        =   {F2}
      End
      Begin VB.Menu About 
         Caption         =   "关于"
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
If MsgBox("此操作会关闭本会话打开的所有标签页，您真的要继续吗？", vbYesNo + vbExclamation) = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub MDIForm_Load()
'切换到IE11内核
CreateObject("wscript.shell").regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", "11000", "REG_DWORD"
'检测命令行参数，如果有则加载命令行参数指定的链接,如果没有则加载主页
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
