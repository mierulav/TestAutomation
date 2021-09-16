Option Explicit

Dim Login : Set Login = Login_Page

Parameter("bResult") = Login.Login(Parameter("Username"), Parameter("Password"))




