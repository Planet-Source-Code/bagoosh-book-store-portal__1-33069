<%

'===============================
' Display Menu Form
'-------------------------------
Sub Menu_Show()
  Dim sFormTitle: sFormTitle = ""

'-------------------------------
' Menu Open Event begin
' Menu Open Event end
'-------------------------------

'-------------------------------
' Set URLs
'-------------------------------
  fldField2 = "Default.asp"
  fldHome = "Default.asp"
  fldReg = "Registration.asp"
  fldShop = "ShoppingCart.asp"
  fldField1 = "Login.asp"
  fldAdmin = "AdminMenu.asp"
'-------------------------------
' Menu Show begin
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Menu BeforeShow Event begin
' Menu BeforeShow Event end
'-------------------------------

'-------------------------------
' Show fields
'-------------------------------
  SetVar "Field2", fldField2
  SetVar "Home", fldHome
  SetVar "Reg", fldReg
  SetVar "Shop", fldShop
  SetVar "Field1", fldField1
  SetVar "Admin", fldAdmin
  Parse "FormMenu", False

'-------------------------------
' Menu Show end
'-------------------------------
End Sub
'===============================

%>