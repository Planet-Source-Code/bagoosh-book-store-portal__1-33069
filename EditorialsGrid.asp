<%
'
'    Filename: EditorialsGrid.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' EditorialsGrid CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' EditorialsGrid CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "EditorialsGrid.asp"
sTemplateFileName = "EditorialsGrid.html"
'===============================


'===============================
' EditorialsGrid PageSecurity begin
CheckSecurity(2)
' EditorialsGrid PageSecurity end
'===============================

'===============================
' EditorialsGrid Open Event begin
' EditorialsGrid Open Event end
'===============================

'===============================
' EditorialsGrid OpenAnyPage Event begin
' EditorialsGrid OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' EditorialsGrid Show begin

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
editorials_Show
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

' EditorialsGrid Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' EditorialsGrid Close Event begin
' EditorialsGrid Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Grid Form
'-------------------------------
Sub editorials_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Editorials"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0
  Dim iPage : iPage = 0
  Dim bEof : bEof = False
  Dim sActionFileName : sActionFileName = "EditorialsRecord.asp"

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  sOrder = " order by e.article_title Asc"
  iSort = GetParam("Formeditorials_Sorting")
  iSorted = GetParam("Formeditorials_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "Formeditorials_Sorting=" & iSort & "&Formeditorials_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "Formeditorials_Sorting=" & iSort & "&Formeditorials_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by e.[article_title]" & sDirection
    if iSort = 2 then sOrder = " order by e1.[editorial_cat_name]" & sDirection
    if iSort = 3 then sOrder = " order by i.[name]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [e].[article_id] as e_article_id, " & _
    "[e].[article_title] as e_article_title, " & _
    "[e].[editorial_cat_id] as e_editorial_cat_id, " & _
    "[e].[item_id] as e_item_id, " & _
    "[e1].[editorial_cat_id] as e1_editorial_cat_id, " & _
    "[e1].[editorial_cat_name] as e1_editorial_cat_name, " & _
    "[i].[item_id] as i_item_id, " & _
    "[i].[name] as i_name " & _
    " from [editorials] e, [editorial_categories] e1, [items] i" & _
    " where [e1].[editorial_cat_id]=e.[editorial_cat_id] and [i].[item_id]=e.[item_id]  "
'-------------------------------

'-------------------------------
' editorials Open Event begin
' editorials Open Event end
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
    SetVar "DListeditorials", ""
    Parse "editorialsNoRecords", False
    SetVar "editorialsNavigator", ""
    Parse "Formeditorials", False
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
  iPage = GetParam("Formeditorials_Page")
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
    fldarticle_id = GetValue(rs, "e_article_id")
    fldarticle_title_URLLink = "EditorialsRecord.asp"
    fldarticle_title_article_id = GetValue(rs, "e_article_id")
    fldarticle_title = GetValue(rs, "e_article_title")
    fldeditorial_cat_id = GetValue(rs, "e1_editorial_cat_name")
    flditem_id = GetValue(rs, "i_name")
'-------------------------------
' editorials Show begin
'-------------------------------

'-------------------------------
' editorials Show Event begin
' editorials Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "article_id", ToHTML(fldarticle_id)
      SetVar "article_title", ToHTML(fldarticle_title)
      SetVar "article_title_URLLink", fldarticle_title_URLLink
      SetVar "Prmarticle_title_article_id", ToURL(fldarticle_title_article_id)
      SetVar "editorial_cat_id", ToHTML(fldeditorial_cat_id)
      SetVar "item_id", ToHTML(flditem_id)
    Parse "DListeditorials", True

'-------------------------------
' editorials Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' editorials Navigation begin
'-------------------------------
  bEof = rs.eof
  if rs.eof and iPage = 1 then
	SetVar "editorialsNavigator", ""
  else
    if bEof then
      SetVar "editorialsNavigatorLastPage", "_"
    else
      SetVar "NextPage", (iPage + 1)
    end if
    if iPage = 1 then
      SetVar "editorialsNavigatorFirstPage", "_"
    else
      SetVar "PrevPage", (iPage - 1)
    end if
    SetVar "editorialsCurrentPage", iPage
    Parse "editorialsNavigator", False
  end if
'-------------------------------
' editorials Navigation end
'-------------------------------

'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "editorialsNoRecords", ""
  Parse "Formeditorials", False

'-------------------------------
' editorials Close Event begin
' editorials Close Event end
'-------------------------------
End Sub
'===============================

%>