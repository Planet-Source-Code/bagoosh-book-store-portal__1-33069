<%
'
'    Filename: AdminMenu.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' AdminMenu CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' AdminMenu CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "AdminMenu.asp"
sTemplateFileName = "AdminMenu.html"
'===============================


'===============================
' AdminMenu PageSecurity begin
CheckSecurity(2)
' AdminMenu PageSecurity end
'===============================

'===============================
' AdminMenu Open Event begin
' AdminMenu Open Event end
'===============================

'===============================
' AdminMenu OpenAnyPage Event begin
' AdminMenu OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' AdminMenu Show begin

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
Form_Show
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

' AdminMenu Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' AdminMenu Close Event begin
' AdminMenu Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Menu Form
'-------------------------------
Sub Form_Show()
  Dim sFormTitle: sFormTitle = "Administration Menu"

'-------------------------------
' Form Open Event begin
' Form Open Event end
'-------------------------------

'-------------------------------
' Set URLs
'-------------------------------
  fldField1 = "MembersGrid.asp"
  fldField2 = "OrdersGrid.asp"
  fldField3 = "AdminBooks.asp"
  fldField4 = "CategoriesGrid.asp"
  fldField5 = "EditorialsGrid.asp"
  fldField6 = "EditorialCatGrid.asp"
  fldField = "CardTypesGrid.asp"
'-------------------------------
' Form Show begin
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Form BeforeShow Event begin
' Form BeforeShow Event end
'-------------------------------

'-------------------------------
' Show fields
'-------------------------------
  SetVar "Field1", fldField1
  SetVar "Field2", fldField2
  SetVar "Field3", fldField3
  SetVar "Field4", fldField4
  SetVar "Field5", fldField5
  SetVar "Field6", fldField6
  SetVar "Field", fldField
  Parse "FormForm", False

'-------------------------------
' Form Show end
'-------------------------------
End Sub
'===============================

%>