<%

'===============================
' Display Menu Form
'-------------------------------
Sub Footer_Show()
  Dim sFormTitle: sFormTitle = ""

'-------------------------------
' Footer Open Event begin
' Footer Open Event end
'-------------------------------

'-------------------------------
' Set URLs
'-------------------------------
  fldField1 = "Default.asp"
  fldField3 = "Registration.asp"
  fldField5 = "ShoppingCart.asp"
  fldField2 = "Login.asp"
  fldField4 = "AdminMenu.asp"
'-------------------------------
' Footer Show begin
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Footer BeforeShow Event begin
' Footer BeforeShow Event end
'-------------------------------

'-------------------------------
' Show fields
'-------------------------------
  SetVar "Field1", fldField1
  SetVar "Field3", fldField3
  SetVar "Field5", fldField5
  SetVar "Field2", fldField2
  SetVar "Field4", fldField4
  Parse "FormFooter", False

'-------------------------------
' Footer Show end
'-------------------------------
End Sub
'===============================

%>