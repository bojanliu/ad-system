Attribute VB_Name = "auto_login"
Option Explicit

Sub Gotosem() '��¼SEM
With CreateObject("internetexplorer.application")
    .Visible = False
    .Navigate "http://192.168.16.156:8080/sem2/login" 'URL
    Do Until .Readystate = 4
    DoEvents
    Loop
        .Document.all("login").Value = "******"
        .Document.all("password").Value = "******"
        .Document.forms(0).submit '����¼��ťδ��������ʹ�ô˷���
        .Quit
End With
End Sub

Sub Gotopro() '��¼PRO
With CreateObject("internetexplorer.application")
    .Visible = False
    .Navigate "http://pro.vemic.com/system/login.php" 'URL
    Do Until .Readystate = 4
    DoEvents
    Loop
    On Error Resume Next
        .Document.all("username").Value = "******"
        .Document.all("password").Value = "******"
        .Document.all("login_btn").Click '��¼��ť�����������ô˷���
        .Quit
End With
End Sub



