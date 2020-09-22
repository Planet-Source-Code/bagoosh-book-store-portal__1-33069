<%
'
'    Filename: CategoriesGrid.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' CategoriesGrid CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' CategoriesGrid CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "CategoriesGrid.asp"
sTemplateFileName = "CategoriesGrid.html"
'===============================


'===============================
' CategoriesGrid PageSecurity begin
CheckSecurity(2)
' CategoriesGrid PageSecurity end
'===============================

'===============================
' CategoriesGrid Open Event begin
' CategoriesGrid Open Event end
'===============================

'===============================
' CategoriesGrid OpenAnyPage Event begin
' CategoriesGrid OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' CategoriesGrid Show begin

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
Categories_Show
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

' CategoriesGrid Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' CategoriesGrid Close Event begin
' CategoriesGrid Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Grid Form
'-------------------------------
Sub Categories_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Categories"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0
  Dim iPage : iPage = 0
  Dim bEof : bEof = False
  Dim sActionFileName : sActionFileName = "CategoriesRecord.asp"

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  sOrder = " order by c.name Asc"
  iSort = GetParam("FormCategories_Sorting")
  iSorted = GetParam("FormCategories_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "FormCategories_Sorting=" & iSort & "&FormCategories_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "FormCategories_Sorting=" & iSort & "&FormCategories_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by c.[name]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [c].[category_id] as c_category_id, " & _
    "[c].[name] as c_name " & _
    " from [categories] c "
'-------------------------------

'-------------------------------
' Categories Open Event begin
' Categories Open Event end
'-------------------------------

'-------------------------------
' Assemble full SQL statement
'-------------------------------
  sSQL = sSQL & sWhere & sOrder
'-------------------------------

SetVar "FormTitle", sFormTitle

'-------------------------------
' Process the link to the record page
'-------------------------------
  SetVar "FormAction", sActionFileName
'-------------------------------

'-------------------------------
' Process the parameters for sorting
'-------------------------------
  SetVar "SortParams", sSortParams
'-------------------------------

'-------------------------------
' Open the recordset
'-------------------------------
  openrs rs, sSQL
'-------------------------------

'-------------------------------
' Process empty recordset
'-------------------------------
  if rs.eof then
    set rs = nothing
    SetVar "DListCategories", ""
    Parse "CategoriesNoRecords", False
    SetVar "CategoriesNavigator", ""
    Parse "FormCategories", False
    exit sub
  end if
'-------------------------------

'-------------------------------
' Initialize page counter and records per page
'-------------------------------
  iRecordsPerPage = 20
  iCounter = 0
'-------------------------------

'-------------------------------
' Process page scroller
'-------------------------------
  iPage = GetParam("FormCategories_Page")
  if IsEmpty(iPage) then iPage = 1 else iPage = CLng(iPage)
  while not rs.eof and iCounter < (iPage-1)*iRecordsPerPage
    rs.movenext
    iCounter = iCounter + 1
  wend
  iCounter = 0
'-------------------------------

'-------------------------------
' Display grid based on recordset
'-------------------------------
  while not rs.EOF  and iCounter < iRecordsPerPage
'-------------------------------
' Create field variables based on database fields
'-------------------------------
    fldname_URLLink = "CategoriesRecord.asp"
    fldname_category_id = GetValue(rs, "c_category_id")
    fldname = GetValue(rs, "c_name")
'-------------------------------
' Categories Show begin
'-------------------------------

'-------------------------------
' Categories Show Event begin
' Categories Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "name", ToHTML(fldname)
      SetVar "name_URLLink", fldname_URLLink
      SetVar "Prmname_category_id", ToURL(fldname_category_id)
    Parse "DListCategories", True

'-------------------------------
' Categories Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' Categories Navigation begin
'-------------------------------
  bEof = rs.eof
  if rs.eof and iPage = 1 then
	SetVar "CategoriesNavigator", ""
  else
    if bEof then
      SetVar "CategoriesNavigatorLastPage", "_"
    else
      SetVar "NextPage", (iPage + 1)
    end if
    if iPage = 1 then
      SetVar "CategoriesNavigatorFirstPage", "_"
    else
      SetVar "PrevPage", (iPage - 1)
    end if
    SetVar "CategoriesCurrentPage", iPage
    Parse "CategoriesNavigator", False
  end if
'-------------------------------
' Categories Navigation end
'-------------------------------

'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "CategoriesNoRecords", ""
  Parse "FormCategories", False

'-------------------------------
' Categories Close Event begin
' Categories Close Event end
'-------------------------------
End Sub
'===============================

%>