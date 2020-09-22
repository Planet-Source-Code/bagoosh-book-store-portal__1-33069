<%
'
'    Filename: Login.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' Login CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' Login CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "Login.asp"
sTemplateFileName = "Login.html"
'===============================


'===============================
' Login PageSecurity begin
' Login PageSecurity end
'===============================

'===============================
' Login Open Event begin
' Login Open Event end
'===============================

'===============================
' Login OpenAnyPage Event begin
' Login OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' Login Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sLoginErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Login"
    LoginAction(sAction)
end select
'===============================

'===============================
' Display page
'-------------------------------
' Load HTML template for this page
'-------------------------------
LoadTemplate sAppPath & sTemplateFileName, "main"
'-------------------------------
' Load HTML template of Header and Footer
'-------------------------------
LoadTemplate sHeaderFileName, "Header"
LoadTemplate sFooterFileName, "Footer"
'-------------------------------
SetVar "FileName", sFileName


'-------------------------------
' Step through each form
'-------------------------------
Menu_Show
Footer_Show
Login_Show
'-------------------------------
' Process page templates
'-------------------------------
Parse "Header", False
Parse "Footer", False
Parse "main", False
'-------------------------------
' Output the page to the browser
'-------------------------------
Response.write PrintVar("main")

' Login Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' Login Close Event begin
' Login Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Login Form Action
'-------------------------------
Sub LoginAction(sAction)
  sQueryString = GetParam("querystring")
  sPage = GetParam("ret_page")
  Select case sAction
    Case "login"

'-------------------------------
' Login Login begin
'-------------------------------
      sLogin = GetParam("Login")
      sPassword = GetParam("Password")
      bPassed = CLng(DLookUp("members", "count(*)", "member_login =" & ToSQL(sLogin, "Text") & " and member_password=" & ToSQL(sPassword, "Text")))

'-------------------------------
' Login OnLogin Event begin
' Login OnLogin Event end
'-------------------------------
      if bPassed > 0 then
'-------------------------------
' Login and password passed
'-------------------------------
        Session("UserID") = CStr(DLookUp("members", "member_id", "member_login =" & ToSQL(sLogin, "Text") & " and member_password=" & ToSQL(sPassword, "Text")))
        Session("UserRights") = CLng(DLookUp("members", "member_level", "member_login =" & ToSQL(sLogin, "Text") & " and member_password=" & ToSQL(sPassword, "Text")))
        cn.Close
        Set cn = Nothing
        if not(sPage = request.serverVariables("SCRIPT_NAME")) and not(isEmpty(sPage)) then
          response.redirect(sPage & "?" & sQueryString)
        end if
        response.redirect("ShoppingCart.asp")
      else
        sLoginErr = "Login or Password is incorrect."
      end if
'-------------------------------
' Login Login end
'-------------------------------
    Case "logout"
'-------------------------------
' Logout action
'-------------------------------
'-------------------------------
' Login Logout begin
'-------------------------------

'-------------------------------
' Login OnLogout Event begin
' Login OnLogout Event end
'-------------------------------
      Session("UserID") = Empty
      Session("UserRights") = Empty
      cn.Close
      Set cn = Nothing
      if not isEmpty(sPage) then response.redirect(sPage & "?" & sQueryString)
      response.redirect(sFileName)
'-------------------------------
' Login Logout end
'-------------------------------
  End Select
End Sub
'===============================

'===============================
' Display Login Form
'-------------------------------
Sub Login_Show()
  Dim sFormTitle: sFormTitle = "Enter login and password"

'-------------------------------
' Login Show begin
'-------------------------------

'-------------------------------
' Login Open Event begin
' Login Open Event end
'-------------------------------
  SetVar "FormTitle", sFormTitle
  SetVar "sLoginErr", sLoginErr
  SetVar "querystring", GetParam("querystring")
  SetVar "ret_page", GetParam("ret_page")
'-------------------------------
' Login BeforeShow Event begin
' Login BeforeShow Event end
'-------------------------------
  if Session("UserID") = "" then
'-------------------------------
' User is not logged in
'-------------------------------
    SetVar "LogoutAct", ""
    SetVar "UserInd", ""
    SetVar "Login", ToHTML(GetParam("Login"))
    if sLoginErr = "" then
      SetVar "LoginError", ""
    else
      SetVar "sLoginErr", sLoginErr
      Parse "LoginError", False
	  End if
    Parse "LoginAct", false
  else
'-------------------------------
' User logged in
'-------------------------------
    SetVar "LoginError", ""
    SetVar "LoginAct", ""
    SetVar "UserID", DLookUp("members", "member_login", "member_id =" & Session("UserID"))
    Parse "UserInd", False
  end if
  Parse "FormLogin", False

'-------------------------------
' Login Close Event begin
' Login Close Event end
'-------------------------------

'-------------------------------
' Login Show end
'-------------------------------
End Sub
'===============================

%>