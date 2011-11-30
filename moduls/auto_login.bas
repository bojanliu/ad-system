Attribute VB_Name = "auto_login"
Option Explicit

Sub Gotosem() '登录SEM
With CreateObject("internetexplorer.application")
    .Visible = False
    .Navigate "http://192.168.16.156:8080/sem2/login" 'URL
    Do Until .Readystate = 4
    DoEvents
    Loop
        .Document.all("login").Value = "liuhuawei"
        .Document.all("password").Value = "adw56557"
        .Document.forms(0).submit '若登录按钮未命名，则使用此方法
        .Quit
End With
End Sub

Sub Gotopro() '登录PRO
With CreateObject("internetexplorer.application")
    .Visible = False
    .Navigate "http://pro.vemic.com/system/login.php" 'URL
    Do Until .Readystate = 4
    DoEvents
    Loop
    On Error Resume Next
        .Document.all("username").Value = "liuhuawei"
        .Document.all("password").Value = "adw56557"
        .Document.all("login_btn").Click '登录按钮有命名，故用此方法
        .Quit
End With
End Sub



