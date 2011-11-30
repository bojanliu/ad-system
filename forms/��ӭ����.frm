VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 欢迎窗体 
   Caption         =   "Ad System"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10365
   OleObjectBlob   =   "欢迎窗体.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "欢迎窗体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
  time.Caption = "5"
  Call b
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '禁止用户通过关闭右上角关闭按钮关闭窗体
    If CloseMode = vbFormControlMenu Then
        MsgBox "您现在无法关闭窗口", , "Ad system"
        Cancel = True
   End If
End Sub
