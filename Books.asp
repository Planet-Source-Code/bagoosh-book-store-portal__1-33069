<%
'
'    Filename: Books.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' Books CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' Books CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "Books.asp"
sTemplateFileName = "Books.html"
'===============================


'===============================
' Books PageSecurity begin
' Books PageSecurity end
'===============================

'===============================
' Books Open Event begin
' Books Open Event end
'===============================

'===============================
' Books OpenAnyPage Event begin
' Books OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' Books Show begin

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
AdvMenu_Show
Total_Show
Results_Show
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

' Books Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' Books Close Event begin
' Books Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Grid Form
'-------------------------------
Sub Results_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Search Results"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0
  Dim iPage : iPage = 0
  Dim bEof : bEof = False

  SetVar "TransitParams", "author=" & ToURL(GetParam("author")) & "&category_id=" & ToURL(GetParam("category_id")) & "&name=" & ToURL(GetParam("name")) & "&pricemax=" & ToURL(GetParam("pricemax")) & "&pricemin=" & ToURL(GetParam("pricemin")) & "&"
  SetVar "FormParams", "author=" & ToURL(GetParam("author")) & "&category_id=" & ToURL(GetParam("category_id")) & "&name=" & ToURL(GetParam("name")) & "&pricemax=" & ToURL(GetParam("pricemax")) & "&pricemin=" & ToURL(GetParam("pricemin")) & "&"

'-------------------------------
' Build WHERE statement
'-------------------------------
  pauthor = GetParam("author")
  if not isEmpty(pauthor) then
    HasParam = true
    sWhere = sWhere & "i.[author] like '%" & replace(pauthor, "'", "''") & "%'"
  end if
  pcategory_id = GetParam("category_id")
  if IsNumeric(pcategory_id) and not isEmpty(pcategory_id) then pcategory_id = ToSQL(pcategory_id, "Number") else pcategory_id = Empty
  if not isEmpty(pcategory_id) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[category_id]=" & pcategory_id
  end if
  pname = GetParam("name")
  if not isEmpty(pname) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[name] like '%" & replace(pname, "'", "''") & "%'"
  end if
  ppricemax = GetParam("pricemax")
  if IsNumeric(ppricemax) and not isEmpty(ppricemax) then ppricemax = ToSQL(ppricemax, "Number") else ppricemax = Empty
  if not isEmpty(ppricemax) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[price]<" & ppricemax
  end if
  ppricemin = GetParam("pricemin")
  if IsNumeric(ppricemin) and not isEmpty(ppricemin) then ppricemin = ToSQL(ppricemin, "Number") else ppricemin = Empty
  if not isEmpty(ppricemin) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[price]>" & ppricemin
  end if


  if HasParam then
    sWhere = " AND (" & sWhere & ")"
  end if
  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  sOrder = " order by i.name Asc"
  iSort = GetParam("FormResults_Sorting")
  iSorted = GetParam("FormResults_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "FormResults_Sorting=" & iSort & "&FormResults_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "FormResults_Sorting=" & iSort & "&FormResults_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by i.[name]" & sDirection
    if iSort = 2 then sOrder = " order by i.[author]" & sDirection
    if iSort = 3 then sOrder = " order by i.[price]" & sDirection
    if iSort = 4 then sOrder = " order by c.[name]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [i].[author] as i_author, " & _
    "[i].[category_id] as i_category_id, " & _
    "[i].[image_url] as i_image_url, " & _
    "[i].[item_id] as i_item_id, " & _
    "[i].[name] as i_name, " & _
    "[i].[price] as i_price, " & _
    "[c].[category_id] as c_category_id, " & _
    "[c].[name] as c_name " & _
    " from [items] i, [categories] c" & _
    " where [c].[category_id]=i.[category_id]  "
'-------------------------------

'-------------------------------
' Results Open Event begin
' Results Open Event end
'-------------------------------

'-------------------------------
' Assemble full SQL statement
'-------------------------------
  sSQL = sSQL & sWhere & sOrder
'-------------------------------

SetVar "FormTitle", sFormTitle

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
    SetVar "DListResults", ""
    Parse "ResultsNoRecords", False
    SetVar "ResultsNavigator", ""
    Parse "FormResults", False
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
  iPage = GetParam("FormResults_Page")
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
    fldauthor = GetValue(rs, "i_author")
    fldcategory_id = GetValue(rs, "c_name")
    fldimage_url = GetValue(rs, "i_image_url")
    fldname_URLLink = "BookDetail.asp"
    fldname_item_id = GetValue(rs, "i_item_id")
    fldname = GetValue(rs, "i_name")
    fldprice = GetValue(rs, "i_price")
'-------------------------------
' Results Show begin
'-------------------------------

'-------------------------------
' Results Show Event begin
fldname="<img border=0 src=" & fldimage_url & "></td><td valign=""top"" width=""100%""><table><tr><td style=""background-color: #FFFFFF; border-style: inset; border-width: 0""><font style=""font-size: 10pt; color: #CE7E00; font-weight: bold""><b>" & fldname & "</b>"
' Results Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "name", fldname
      SetVar "name_URLLink", fldname_URLLink
      SetVar "Prmname_item_id", ToURL(fldname_item_id)
      SetVar "author", ToHTML(fldauthor)
      SetVar "price", ToHTML(fldprice)
      SetVar "category_id", ToHTML(fldcategory_id)
      SetVar "image_url", ToHTML(fldimage_url)
'-------------------------------
' Process the record separator
'-------------------------------
    if rs.EOF or iCounter = iRecordsPerPage-1 then
       SetVar "ResultsRecordSeparator", ""
    else
      Parse "ResultsRecordSeparator", false
    end if
'-------------------------------
    Parse "DListResults", True

'-------------------------------
' Results Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' Results Navigation begin
'-------------------------------
  bEof = rs.eof
  if rs.eof and iPage = 1 then
	SetVar "ResultsNavigator", ""
  else
    if bEof then
      SetVar "ResultsNavigatorLastPage", "_"
    else
      SetVar "NextPage", (iPage + 1)
    end if
    if iPage = 1 then
      SetVar "ResultsNavigatorFirstPage", "_"
    else
      SetVar "PrevPage", (iPage - 1)
    end if
    SetVar "ResultsCurrentPage", iPage
    Parse "ResultsNavigator", False
  end if
'-------------------------------
' Results Navigation end
'-------------------------------

'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "ResultsNoRecords", ""
  Parse "FormResults", False

'-------------------------------
' Results Close Event begin
' Results Close Event end
'-------------------------------
End Sub
'===============================


'===============================
' Display Search Form
'-------------------------------
Sub Search_Show()
  Dim sFormTitle: sFormTitle = ""
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
      fldcategory_id = GetParam("category_id")
      fldname = GetParam("name")

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
    
      SetVar "name", ToHTML(fldname)

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
' Display Menu Form
'-------------------------------
Sub AdvMenu_Show()
  Dim sFormTitle: sFormTitle = ""

'-------------------------------
' AdvMenu Open Event begin
' AdvMenu Open Event end
'-------------------------------

'-------------------------------
' Set URLs
'-------------------------------
  fldField1 = "AdvSearch.asp"
'-------------------------------
' AdvMenu Show begin
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' AdvMenu BeforeShow Event begin
' AdvMenu BeforeShow Event end
'-------------------------------

'-------------------------------
' Show fields
'-------------------------------
  SetVar "Field1", fldField1
  Parse "FormAdvMenu", False

'-------------------------------
' AdvMenu Show end
'-------------------------------
End Sub
'===============================


'===============================
' Display Grid Form
'-------------------------------
Sub Total_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = ""
  Dim HasParam : HasParam = false
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", ""
  SetVar "FormParams", "author=" & ToURL(GetParam("author")) & "&category_id=" & ToURL(GetParam("category_id")) & "&name=" & ToURL(GetParam("name")) & "&pricemax=" & ToURL(GetParam("pricemax")) & "&pricemin=" & ToURL(GetParam("pricemin")) & "&"

'-------------------------------
' Build WHERE statement
'-------------------------------
  pauthor = GetParam("author")
  if not isEmpty(pauthor) then
    HasParam = true
    sWhere = sWhere & "i.[author] like '%" & replace(pauthor, "'", "''") & "%'"
  end if
  pcategory_id = GetParam("category_id")
  if IsNumeric(pcategory_id) and not isEmpty(pcategory_id) then pcategory_id = ToSQL(pcategory_id, "Number") else pcategory_id = Empty
  if not isEmpty(pcategory_id) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[category_id]=" & pcategory_id
  end if
  pname = GetParam("name")
  if not isEmpty(pname) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[name] like '%" & replace(pname, "'", "''") & "%'"
  end if
  ppricemax = GetParam("pricemax")
  if IsNumeric(ppricemax) and not isEmpty(ppricemax) then ppricemax = ToSQL(ppricemax, "Number") else ppricemax = Empty
  if not isEmpty(ppricemax) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[price]<=" & ppricemax
  end if
  ppricemin = GetParam("pricemin")
  if IsNumeric(ppricemin) and not isEmpty(ppricemin) then ppricemin = ToSQL(ppricemin, "Number") else ppricemin = Empty
  if not isEmpty(ppricemin) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "i.[price]>=" & ppricemin
  end if


  if HasParam then
    sWhere = " WHERE (" & sWhere & ")"
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [i].[author] as i_author, " & _
    "[i].[category_id] as i_category_id, " & _
    "[i].[item_id] as i_item_id, " & _
    "[i].[name] as i_name, " & _
    "[i].[price] as i_price " & _
    " from [items] i "
'-------------------------------

'-------------------------------
' Total Open Event begin
sSQL="select count(item_id) as i_item_id from items as i"
' Total Open Event end
'-------------------------------

'-------------------------------
' Assemble full SQL statement
'-------------------------------
  sSQL = sSQL & sWhere & sOrder
'-------------------------------

SetVar "FormTitle", sFormTitle

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
    SetVar "DListTotal", ""
    Parse "TotalNoRecords", False
    Parse "FormTotal", False
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
' Display grid based on recordset
'-------------------------------
  while not rs.EOF  and iCounter < iRecordsPerPage
'-------------------------------
' Create field variables based on database fields
'-------------------------------
    flditem_id = GetValue(rs, "i_item_id")
'-------------------------------
' Total Show begin
'-------------------------------

'-------------------------------
' Total Show Event begin
' Total Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "item_id", ToHTML(flditem_id)
'-------------------------------
' Process the record separator
'-------------------------------
    if rs.EOF or iCounter = iRecordsPerPage-1 then
       SetVar "TotalRecordSeparator", ""
    else
      Parse "TotalRecordSeparator", false
    end if
'-------------------------------
    Parse "DListTotal", True

'-------------------------------
' Total Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "TotalNoRecords", ""
  Parse "FormTotal", False

'-------------------------------
' Total Close Event begin
' Total Close Event end
'-------------------------------
End Sub
'===============================

%>