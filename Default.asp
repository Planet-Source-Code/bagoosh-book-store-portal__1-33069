<%
'
'    Filename: Default.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' Default CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' Default CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "Default.asp"
sTemplateFileName = "Default.html"
'===============================


'===============================
' Default PageSecurity begin
' Default PageSecurity end
'===============================

'===============================
' Default Open Event begin
' Default Open Event end
'===============================

'===============================
' Default OpenAnyPage Event begin
' Default OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' Default Show begin

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
Categories_Show
Specials_Show
Recommended_Show
What_Show
New_Show
Weekly_Show
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

' Default Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' Default Close Event begin
' Default Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Search Form
'-------------------------------
Sub Search_Show()
  Dim sFormTitle: sFormTitle = "Search"
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
  Dim sFormTitle: sFormTitle = "More Search Options"

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
Sub Recommended_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Recommended Titles"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 5
  Dim iCounter : iCounter = 0
  Dim iPage : iPage = 0
  Dim bEof : bEof = False

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  sWhere = " WHERE is_recommended=1"
  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  iSort = GetParam("FormRecommended_Sorting")
  iSorted = GetParam("FormRecommended_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "FormRecommended_Sorting=" & iSort & "&FormRecommended_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "FormRecommended_Sorting=" & iSort & "&FormRecommended_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by i.[name]" & sDirection
    if iSort = 2 then sOrder = " order by i.[author]" & sDirection
    if iSort = 3 then sOrder = " order by i.[price]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [i].[author] as i_author, " & _
    "[i].[image_url] as i_image_url, " & _
    "[i].[item_id] as i_item_id, " & _
    "[i].[name] as i_name, " & _
    "[i].[price] as i_price " & _
    " from [items] i "
'-------------------------------

'-------------------------------
' Recommended Open Event begin
' Recommended Open Event end
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
    SetVar "DListRecommended", ""
    Parse "RecommendedNoRecords", False
    SetVar "RecommendedNavigator", ""
    Parse "FormRecommended", False
    exit sub
  end if
'-------------------------------

'-------------------------------
' Initialize page counter and records per page
'-------------------------------
  iRecordsPerPage = 5
  iCounter = 0
'-------------------------------

'-------------------------------
' Process page scroller
'-------------------------------
  iPage = GetParam("FormRecommended_Page")
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
    fldimage_url = GetValue(rs, "i_image_url")
    fldname_URLLink = "BookDetail.asp"
    fldname_item_id = GetValue(rs, "i_item_id")
    fldname = GetValue(rs, "i_name")
    fldprice = GetValue(rs, "i_price")
'-------------------------------
' Recommended Show begin
'-------------------------------

'-------------------------------
' Recommended Show Event begin
fldname="<img border=""0"" src=""" & fldimage_url & """></td><td valign=""top""><table width=""100%"" style=""width:100%""><tr><td style=""background-color: #FFFFFF; border-style: inset; border-width: 0""><font style=""font-size: 10pt; color: #CE7E00; font-weight: bold""><b>" & fldname & "</b>"
' Recommended Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "name", fldname
      SetVar "name_URLLink", fldname_URLLink
      SetVar "Prmname_item_id", ToURL(fldname_item_id)
      SetVar "author", ToHTML(fldauthor)
      SetVar "image_url", ToHTML(fldimage_url)
      SetVar "price", ToHTML(fldprice)
'-------------------------------
' Process the record separator
'-------------------------------
    if rs.EOF or iCounter = iRecordsPerPage-1 then
       SetVar "RecommendedRecordSeparator", ""
    else
      Parse "RecommendedRecordSeparator", false
    end if
'-------------------------------
    Parse "DListRecommended", True

'-------------------------------
' Recommended Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' Recommended Navigation begin
'-------------------------------
  bEof = rs.eof
  if rs.eof and iPage = 1 then
	SetVar "RecommendedNavigator", ""
  else
    if bEof then
      SetVar "RecommendedNavigatorLastPage", "_"
    else
      SetVar "NextPage", (iPage + 1)
    end if
    if iPage = 1 then
      SetVar "RecommendedNavigatorFirstPage", "_"
    else
      SetVar "PrevPage", (iPage - 1)
    end if
    SetVar "RecommendedCurrentPage", iPage
    Parse "RecommendedNavigator", False
  end if
'-------------------------------
' Recommended Navigation end
'-------------------------------

'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "RecommendedNoRecords", ""
  Parse "FormRecommended", False

'-------------------------------
' Recommended Close Event begin
' Recommended Close Event end
'-------------------------------
End Sub
'===============================


'===============================
' Display Grid Form
'-------------------------------
Sub What_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "What We're Reading"
  Dim HasParam : HasParam = false
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  sWhere = " WHERE editorial_cat_id=1"

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [e].[article_desc] as e_article_desc, " & _
    "[e].[article_title] as e_article_title, " & _
    "[e].[item_id] as e_item_id " & _
    " from [editorials] e "
'-------------------------------

'-------------------------------
' What Open Event begin
' What Open Event end
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
    SetVar "DListWhat", ""
    Parse "WhatNoRecords", False
    Parse "FormWhat", False
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
    fldarticle_desc = GetValue(rs, "e_article_desc")
    fldarticle_title_URLLink = "BookDetail.asp"
    fldarticle_title_item_id = GetValue(rs, "e_item_id")
    fldarticle_title = GetValue(rs, "e_article_title")
    flditem_id = GetValue(rs, "e_item_id")
'-------------------------------
' What Show begin
'-------------------------------

'-------------------------------
' What Show Event begin
fldarticle_title="<b>" & fldarticle_title & "</b>"
fldarticle_desc="<img align=""left"" border=""0"" src=""" & dlookup("items","image_url","item_id=" & flditem_id) & """>" & fldarticle_desc
' What Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "article_title", fldarticle_title
      SetVar "article_title_URLLink", fldarticle_title_URLLink
      SetVar "Prmarticle_title_item_id", ToURL(fldarticle_title_item_id)
      SetVar "article_desc", fldarticle_desc
      SetVar "item_id", ToHTML(flditem_id)
'-------------------------------
' Process the record separator
'-------------------------------
    if rs.EOF or iCounter = iRecordsPerPage-1 then
       SetVar "WhatRecordSeparator", ""
    else
      Parse "WhatRecordSeparator", false
    end if
'-------------------------------
    Parse "DListWhat", True

'-------------------------------
' What Show end
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
  SetVar "WhatNoRecords", ""
  Parse "FormWhat", False

'-------------------------------
' What Close Event begin
' What Close Event end
'-------------------------------
End Sub
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
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", ""
  SetVar "FormParams", ""



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
' Display grid based on recordset
'-------------------------------
  while not rs.EOF  and iCounter < iRecordsPerPage
'-------------------------------
' Create field variables based on database fields
'-------------------------------
    fldname_URLLink = "Books.asp"
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


'===============================
' Display Grid Form
'-------------------------------
Sub New_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "New & Notable"
  Dim HasParam : HasParam = false
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  sWhere = " WHERE editorial_cat_id=2"

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [e].[article_desc] as e_article_desc, " & _
    "[e].[article_title] as e_article_title, " & _
    "[e].[item_id] as e_item_id " & _
    " from [editorials] e "
'-------------------------------

'-------------------------------
' New Open Event begin
' New Open Event end
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
    SetVar "DListNew", ""
    Parse "NewNoRecords", False
    Parse "FormNew", False
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
    fldarticle_desc = GetValue(rs, "e_article_desc")
    fldarticle_title_URLLink = "BookDetail.asp"
    fldarticle_title_item_id = GetValue(rs, "e_item_id")
    fldarticle_title = GetValue(rs, "e_article_title")
    flditem_id = GetValue(rs, "e_item_id")
'-------------------------------
' New Show begin
'-------------------------------

'-------------------------------
' New Show Event begin
fldarticle_title="<b>" & fldarticle_title & "</b>"
fldarticle_desc="<img align=""left"" border=""0"" src=""" & dlookup("items","image_url","item_id=" & flditem_id) & """>" & fldarticle_desc
' New Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "article_title", fldarticle_title
      SetVar "article_title_URLLink", fldarticle_title_URLLink
      SetVar "Prmarticle_title_item_id", ToURL(fldarticle_title_item_id)
      SetVar "article_desc", fldarticle_desc
      SetVar "item_id", ToHTML(flditem_id)
'-------------------------------
' Process the record separator
'-------------------------------
    if rs.EOF or iCounter = iRecordsPerPage-1 then
       SetVar "NewRecordSeparator", ""
    else
      Parse "NewRecordSeparator", false
    end if
'-------------------------------
    Parse "DListNew", True

'-------------------------------
' New Show end
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
  SetVar "NewNoRecords", ""
  Parse "FormNew", False

'-------------------------------
' New Close Event begin
' New Close Event end
'-------------------------------
End Sub
'===============================


'===============================
' Display Grid Form
'-------------------------------
Sub Weekly_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "This Week's Featured Books"
  Dim HasParam : HasParam = false
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  sWhere = " WHERE editorial_cat_id=3"

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [e].[article_desc] as e_article_desc, " & _
    "[e].[article_title] as e_article_title, " & _
    "[e].[item_id] as e_item_id " & _
    " from [editorials] e "
'-------------------------------

'-------------------------------
' Weekly Open Event begin
' Weekly Open Event end
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
    SetVar "DListWeekly", ""
    Parse "WeeklyNoRecords", False
    Parse "FormWeekly", False
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
    fldarticle_desc = GetValue(rs, "e_article_desc")
    fldarticle_title_URLLink = "BookDetail.asp"
    fldarticle_title_item_id = GetValue(rs, "e_item_id")
    fldarticle_title = GetValue(rs, "e_article_title")
    flditem_id = GetValue(rs, "e_item_id")
'-------------------------------
' Weekly Show begin
'-------------------------------

'-------------------------------
' Weekly Show Event begin
fldarticle_title="<b>" & fldarticle_title & "</b>"
fldarticle_desc="<img align=""left"" border=""0"" src=""" & dlookup("items","image_url","item_id=" & flditem_id) & """>" & fldarticle_desc
' Weekly Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "article_title", fldarticle_title
      SetVar "article_title_URLLink", fldarticle_title_URLLink
      SetVar "Prmarticle_title_item_id", ToURL(fldarticle_title_item_id)
      SetVar "article_desc", fldarticle_desc
      SetVar "item_id", ToHTML(flditem_id)
'-------------------------------
' Process the record separator
'-------------------------------
    if rs.EOF or iCounter = iRecordsPerPage-1 then
       SetVar "WeeklyRecordSeparator", ""
    else
      Parse "WeeklyRecordSeparator", false
    end if
'-------------------------------
    Parse "DListWeekly", True

'-------------------------------
' Weekly Show end
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
  SetVar "WeeklyNoRecords", ""
  Parse "FormWeekly", False

'-------------------------------
' Weekly Close Event begin
' Weekly Close Event end
'-------------------------------
End Sub
'===============================


'===============================
' Display Grid Form
'-------------------------------
Sub Specials_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Weekly Specials"
  Dim HasParam : HasParam = false
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  sWhere = " WHERE editorial_cat_id=4"

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [e].[article_desc] as e_article_desc, " & _
    "[e].[article_title] as e_article_title " & _
    " from [editorials] e "
'-------------------------------

'-------------------------------
' Specials Open Event begin
' Specials Open Event end
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
    SetVar "DListSpecials", ""
    Parse "SpecialsNoRecords", False
    Parse "FormSpecials", False
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
    fldarticle_desc = GetValue(rs, "e_article_desc")
    fldarticle_title = GetValue(rs, "e_article_title")
'-------------------------------
' Specials Show begin
'-------------------------------

'-------------------------------
' Specials Show Event begin
' Specials Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "article_title", fldarticle_title
      SetVar "article_desc", fldarticle_desc
'-------------------------------
' Process the record separator
'-------------------------------
    if rs.EOF or iCounter = iRecordsPerPage-1 then
       SetVar "SpecialsRecordSeparator", ""
    else
      Parse "SpecialsRecordSeparator", false
    end if
'-------------------------------
    Parse "DListSpecials", True

'-------------------------------
' Specials Show end
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
  SetVar "SpecialsNoRecords", ""
  Parse "FormSpecials", False

'-------------------------------
' Specials Close Event begin
' Specials Close Event end
'-------------------------------
End Sub
'===============================

%>