<%
'
'    Filename: AdminBooks.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' AdminBooks CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' AdminBooks CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "AdminBooks.asp"
sTemplateFileName = "AdminBooks.html"
'===============================


'===============================
' AdminBooks PageSecurity begin
CheckSecurity(2)
' AdminBooks PageSecurity end
'===============================

'===============================
' AdminBooks Open Event begin
' AdminBooks Open Event end
'===============================

'===============================
' AdminBooks OpenAnyPage Event begin
' AdminBooks OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' AdminBooks Show begin

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
Items_Show
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

' AdminBooks Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' AdminBooks Close Event begin
' AdminBooks Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Search Form
'-------------------------------
Sub Search_Show()
  Dim sFormTitle: sFormTitle = ""
  Dim sActionFileName: sActionFileName = "AdminBooks.asp"
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
      fldcategory_id = GetParam("category_id")
      fldis_recommended = GetParam("is_recommended")

'-------------------------------
' Search Show begin
'-------------------------------


'-------------------------------
' Search Show Event begin
' Search Show Event end
'-------------------------------
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
    
      SetVar "SearchLBis_recommended", ""
      LOV = Split(";All;0;No;1;Yes", ";")
      if (ubound(LOV) mod 2) = 1 then
        for i = 0 to ubound(LOV) step 2
          SetVar "ID", LOV(i) : SetVar "Value", LOV(i+1)
          if cstr(LOV(i)) = cstr(fldis_recommended) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
          Parse "SearchLBis_recommended", True
        next
      end if
    

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


'===============================
' Display Grid Form
'-------------------------------
Sub Items_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Books"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0
  Dim iPage : iPage = 0
  Dim bEof : bEof = False
  Dim sActionFileName : sActionFileName = "BookMaint.asp"

  SetVar "TransitParams", "category_id=" & ToURL(GetParam("category_id")) & "&is_recommended=" & ToURL(GetParam("is_recommended")) & "&"
  SetVar "FormParams", "category_id=" & ToURL(GetParam("category_id")) & "&is_recommended=" & ToURL(GetParam("is_recommended")) & "&"

'-------------------------------
' Build WHERE statement
'-------------------------------
  pcategory_id = GetParam("category_id")
  if IsNumeric(pcategory_id) and not isEmpty(pcategory_id) then pcategory_id = ToSQL(pcategory_id, "Number") else pcategory_id = Empty
  if not isEmpty(pcategory_id) then
    HasParam = true
    sWhere = sWhere & "i.[category_id]=" & pcategory_id
  end if
  pis_recommended = GetParam("is_recommended")
  if IsNumeric(pis_recommended) and not isEmpty(pis_recommended) then pis_recommended = ToSQL(pis_recommended, "Number") else pis_recommended = Empty
  if not isEmpty(pis_recommended) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[is_recommended]=" & pis_recommended
  end if


  if HasParam then
    sWhere = " AND (" & sWhere & ")"
  end if
  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  iSort = GetParam("FormItems_Sorting")
  iSorted = GetParam("FormItems_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "FormItems_Sorting=" & iSort & "&FormItems_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "FormItems_Sorting=" & iSort & "&FormItems_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by i.[name]" & sDirection
    if iSort = 2 then sOrder = " order by i.[author]" & sDirection
    if iSort = 3 then sOrder = " order by i.[price]" & sDirection
    if iSort = 4 then sOrder = " order by c.[name]" & sDirection
    if iSort = 5 then sOrder = " order by i.[is_recommended]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [i].[author] as i_author, " & _
    "[i].[category_id] as i_category_id, " & _
    "[i].[is_recommended] as i_is_recommended, " & _
    "[i].[item_id] as i_item_id, " & _
    "[i].[name] as i_name, " & _
    "[i].[price] as i_price, " & _
    "[c].[category_id] as c_category_id, " & _
    "[c].[name] as c_name " & _
    " from [items] i, [categories] c" & _
    " where [c].[category_id]=i.[category_id]  "
'-------------------------------

'-------------------------------
' Items Open Event begin
' Items Open Event end
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
    SetVar "DListItems", ""
    Parse "ItemsNoRecords", False
    SetVar "ItemsNavigator", ""
    Parse "FormItems", False
    exit sub
  end if
'-------------------------------

'-------------------------------
' Prepare the lists of values
'-------------------------------

  ais_recommended = Split("0;No;1;Yes", ";")
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
  iPage = GetParam("FormItems_Page")
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
    fldField1_URLLink = "BookMaint.asp"
    fldField1_item_id = GetValue(rs, "i_item_id")
    fldauthor = GetValue(rs, "i_author")
    fldcategory_id = GetValue(rs, "c_name")
    fldis_recommended = GetValue(rs, "i_is_recommended")
    fldname = GetValue(rs, "i_name")
    fldprice = GetValue(rs, "i_price")
    fldField1= "Edit"
'-------------------------------
' Items Show begin
'-------------------------------

'-------------------------------
' Items Show Event begin
' Items Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "Field1", ToHTML(fldField1)
      SetVar "Field1_URLLink", fldField1_URLLink
      SetVar "PrmField1_item_id", ToURL(fldField1_item_id)
      SetVar "name", ToHTML(fldname)
      SetVar "author", ToHTML(fldauthor)
      SetVar "price", ToHTML(fldprice)
      SetVar "category_id", ToHTML(fldcategory_id)
      fldis_recommended = getValFromLOV(fldis_recommended, ais_recommended)
      SetVar "is_recommended", ToHTML(fldis_recommended)
    Parse "DListItems", True

'-------------------------------
' Items Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' Items Navigation begin
'-------------------------------
  bEof = rs.eof
  if rs.eof and iPage = 1 then
	SetVar "ItemsNavigator", ""
  else
    if bEof then
      SetVar "ItemsNavigatorLastPage", "_"
    else
      SetVar "NextPage", (iPage + 1)
    end if
    if iPage = 1 then
      SetVar "ItemsNavigatorFirstPage", "_"
    else
      SetVar "PrevPage", (iPage - 1)
    end if
    SetVar "ItemsCurrentPage", iPage
    Parse "ItemsNavigator", False
  end if
'-------------------------------
' Items Navigation end
'-------------------------------

'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "ItemsNoRecords", ""
  Parse "FormItems", False

'-------------------------------
' Items Close Event begin
' Items Close Event end
'-------------------------------
End Sub
'===============================

%>