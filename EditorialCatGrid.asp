<%
'
'    Filename: EditorialCatGrid.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' EditorialCatGrid CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' EditorialCatGrid CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "EditorialCatGrid.asp"
sTemplateFileName = "EditorialCatGrid.html"
'===============================


'===============================
' EditorialCatGrid PageSecurity begin
CheckSecurity(2)
' EditorialCatGrid PageSecurity end
'===============================

'===============================
' EditorialCatGrid Open Event begin
' EditorialCatGrid Open Event end
'===============================

'===============================
' EditorialCatGrid OpenAnyPage Event begin
' EditorialCatGrid OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' EditorialCatGrid Show begin

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
editorial_categories_Show
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

' EditorialCatGrid Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' EditorialCatGrid Close Event begin
' EditorialCatGrid Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Grid Form
'-------------------------------
Sub editorial_categories_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Editorial Category"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0
  Dim iPage : iPage = 0
  Dim bEof : bEof = False
  Dim sActionFileName : sActionFileName = "EditorialCatRecord.asp"

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  sOrder = " order by e.editorial_cat_name Asc"
  iSort = GetParam("Formeditorial_categories_Sorting")
  iSorted = GetParam("Formeditorial_categories_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "Formeditorial_categories_Sorting=" & iSort & "&Formeditorial_categories_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "Formeditorial_categories_Sorting=" & iSort & "&Formeditorial_categories_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by e.[editorial_cat_name]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [e].[editorial_cat_id] as e_editorial_cat_id, " & _
    "[e].[editorial_cat_name] as e_editorial_cat_name " & _
    " from [editorial_categories] e "
'-------------------------------

'-------------------------------
' editorial_categories Open Event begin
' editorial_categories Open Event end
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
    SetVar "DListeditorial_categories", ""
    Parse "editorial_categoriesNoRecords", False
    SetVar "editorial_categoriesNavigator", ""
    Parse "Formeditorial_categories", False
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
  iPage = GetParam("Formeditorial_categories_Page")
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
    fldeditorial_cat_id = GetValue(rs, "e_editorial_cat_id")
    fldeditorial_cat_name_URLLink = "EditorialCatRecord.asp"
    fldeditorial_cat_name_editorial_cat_id = GetValue(rs, "e_editorial_cat_id")
    fldeditorial_cat_name = GetValue(rs, "e_editorial_cat_name")
'-------------------------------
' editorial_categories Show begin
'-------------------------------

'-------------------------------
' editorial_categories Show Event begin
' editorial_categories Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "editorial_cat_id", ToHTML(fldeditorial_cat_id)
      SetVar "editorial_cat_name", ToHTML(fldeditorial_cat_name)
      SetVar "editorial_cat_name_URLLink", fldeditorial_cat_name_URLLink
      SetVar "Prmeditorial_cat_name_editorial_cat_id", ToURL(fldeditorial_cat_name_editorial_cat_id)
    Parse "DListeditorial_categories", True

'-------------------------------
' editorial_categories Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' editorial_categories Navigation begin
'-------------------------------
  bEof = rs.eof
  if rs.eof and iPage = 1 then
	SetVar "editorial_categoriesNavigator", ""
  else
    if bEof then
      SetVar "editorial_categoriesNavigatorLastPage", "_"
    else
      SetVar "NextPage", (iPage + 1)
    end if
    if iPage = 1 then
      SetVar "editorial_categoriesNavigatorFirstPage", "_"
    else
      SetVar "PrevPage", (iPage - 1)
    end if
    SetVar "editorial_categoriesCurrentPage", iPage
    Parse "editorial_categoriesNavigator", False
  end if
'-------------------------------
' editorial_categories Navigation end
'-------------------------------

'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "editorial_categoriesNoRecords", ""
  Parse "Formeditorial_categories", False

'-------------------------------
' editorial_categories Close Event begin
' editorial_categories Close Event end
'-------------------------------
End Sub
'===============================

%>