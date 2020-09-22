<%
'
'    Filename: AdvSearch.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' AdvSearch CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' AdvSearch CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "AdvSearch.asp"
sTemplateFileName = "AdvSearch.html"
'===============================


'===============================
' AdvSearch PageSecurity begin
' AdvSearch PageSecurity end
'===============================

'===============================
' AdvSearch Open Event begin
' AdvSearch Open Event end
'===============================

'===============================
' AdvSearch OpenAnyPage Event begin
' AdvSearch OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' AdvSearch Show begin

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
Search_Show
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

' AdvSearch Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' AdvSearch Close Event begin
' AdvSearch Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Search Form
'-------------------------------
Sub Search_Show()
  Dim sFormTitle: sFormTitle = "Advanced Search"
  Dim sActionFileName: sActionFileName = "Books.asp"
  Dim scategory_idDisplayValue: scategory_idDisplayValue = "All"

'-------------------------------
' Search Open Event begin
' Search Open Event end
'-------------------------------
      SetVar "FormTitle", sFormTitle
      SetVar "ActionPage", sActionFileName

'-------------------------------
' Set variables with search parameters
'-------------------------------
      fldname = GetParam("name")
      fldauthor = GetParam("author")
      fldcategory_id = GetParam("category_id")
      fldpricemin = GetParam("pricemin")
      fldpricemax = GetParam("pricemax")

'-------------------------------
' Search Show begin
'-------------------------------


'-------------------------------
' Search Show Event begin
' Search Show Event end
'-------------------------------
      SetVar "name", ToHTML(fldname)
      SetVar "author", ToHTML(fldauthor)
      SetVar "SearchLBcategory_id", ""
      SetVar "Selected", ""
      SetVar "ID", ""
      SetVar "Value", scategory_idDisplayValue
      Parse "SearchLBcategory_id", True
      openrs rscategory_id, "select category_id, name from categories order by 2"
      while not rscategory_id.EOF
        SetVar "ID", GetValue(rscategory_id, 0) : SetVar "Value", GetValue(rscategory_id, 1)
        if cstr(GetValue(rscategory_id, 0)) = cstr(fldcategory_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "SearchLBcategory_id", True
        rscategory_id.MoveNext
      wend
      set rscategory_id = nothing
    
      SetVar "pricemin", ToHTML(fldpricemin)
      SetVar "pricemax", ToHTML(fldpricemax)

'-------------------------------
' Search Show end
'-------------------------------

'-------------------------------
' Search Close Event begin
' Search Close Event end
'-------------------------------
      Parse "FormSearch", False
End Sub
'===============================

%>