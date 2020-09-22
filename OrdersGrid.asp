<%
'
'    Filename: OrdersGrid.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' OrdersGrid CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' OrdersGrid CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "OrdersGrid.asp"
sTemplateFileName = "OrdersGrid.html"
'===============================


'===============================
' OrdersGrid PageSecurity begin
CheckSecurity(2)
' OrdersGrid PageSecurity end
'===============================

'===============================
' OrdersGrid Open Event begin
' OrdersGrid Open Event end
'===============================

'===============================
' OrdersGrid OpenAnyPage Event begin
' OrdersGrid OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' OrdersGrid Show begin

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
Orders_Show
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

' OrdersGrid Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' OrdersGrid Close Event begin
' OrdersGrid Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Search Form
'-------------------------------
Sub Search_Show()
  Dim sFormTitle: sFormTitle = ""
  Dim sActionFileName: sActionFileName = "OrdersGrid.asp"
  Dim sitem_idDisplayValue: sitem_idDisplayValue = "All"
  Dim smember_idDisplayValue: smember_idDisplayValue = "All"

'-------------------------------
' Search Open Event begin
' Search Open Event end
'-------------------------------
      SetVar "FormTitle", sFormTitle
      SetVar "ActionPage", sActionFileName

'-------------------------------
' Set variables with search parameters
'-------------------------------
      flditem_id = GetParam("item_id")
      fldmember_id = GetParam("member_id")

'-------------------------------
' Search Show begin
'-------------------------------


'-------------------------------
' Search Show Event begin
' Search Show Event end
'-------------------------------
      SetVar "SearchLBitem_id", ""
      SetVar "Selected", ""
      SetVar "ID", ""
      SetVar "Value", sitem_idDisplayValue
      Parse "SearchLBitem_id", True
      openrs rsitem_id, "select item_id, name from items order by 2"
      while not rsitem_id.EOF
        SetVar "ID", GetValue(rsitem_id, 0) : SetVar "Value", GetValue(rsitem_id, 1)
        if cstr(GetValue(rsitem_id, 0)) = cstr(flditem_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "SearchLBitem_id", True
        rsitem_id.MoveNext
      wend
      set rsitem_id = nothing
    
      SetVar "SearchLBmember_id", ""
      SetVar "Selected", ""
      SetVar "ID", ""
      SetVar "Value", smember_idDisplayValue
      Parse "SearchLBmember_id", True
      openrs rsmember_id, "select member_id, member_login from members order by 2"
      while not rsmember_id.EOF
        SetVar "ID", GetValue(rsmember_id, 0) : SetVar "Value", GetValue(rsmember_id, 1)
        if cstr(GetValue(rsmember_id, 0)) = cstr(fldmember_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "SearchLBmember_id", True
        rsmember_id.MoveNext
      wend
      set rsmember_id = nothing
    

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
Sub Orders_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Orders"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0
  Dim iPage : iPage = 0
  Dim bEof : bEof = False
  Dim sActionFileName : sActionFileName = "OrdersRecord.asp"

  SetVar "TransitParams", "item_id=" & ToURL(GetParam("item_id")) & "&member_id=" & ToURL(GetParam("member_id")) & "&"
  SetVar "FormParams", "item_id=" & ToURL(GetParam("item_id")) & "&member_id=" & ToURL(GetParam("member_id")) & "&"

'-------------------------------
' Build WHERE statement
'-------------------------------
  pitem_id = GetParam("item_id")
  if IsNumeric(pitem_id) and not isEmpty(pitem_id) then pitem_id = ToSQL(pitem_id, "Number") else pitem_id = Empty
  if not isEmpty(pitem_id) then
    HasParam = true
    sWhere = sWhere & "o.[item_id]=" & pitem_id
  end if
  pmember_id = GetParam("member_id")
  if IsNumeric(pmember_id) and not isEmpty(pmember_id) then pmember_id = ToSQL(pmember_id, "Number") else pmember_id = Empty
  if not isEmpty(pmember_id) then
    if not (sWhere = "") then sWhere = sWhere & " and "
    HasParam = true
    sWhere = sWhere & "o.[member_id]=" & pmember_id
  end if


  if HasParam then
    sWhere = " AND (" & sWhere & ")"
  end if
  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  sOrder = " order by o.order_id Asc"
  iSort = GetParam("FormOrders_Sorting")
  iSorted = GetParam("FormOrders_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "FormOrders_Sorting=" & iSort & "&FormOrders_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "FormOrders_Sorting=" & iSort & "&FormOrders_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by i.[name]" & sDirection
    if iSort = 2 then sOrder = " order by m.[member_login]" & sDirection
    if iSort = 3 then sOrder = " order by o.[quantity]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [o].[item_id] as o_item_id, " & _
    "[o].[member_id] as o_member_id, " & _
    "[o].[order_id] as o_order_id, " & _
    "[o].[quantity] as o_quantity, " & _
    "[i].[item_id] as i_item_id, " & _
    "[i].[name] as i_name, " & _
    "[m].[member_id] as m_member_id, " & _
    "[m].[member_login] as m_member_login " & _
    " from [orders] o, [items] i, [members] m" & _
    " where [i].[item_id]=o.[item_id] and [m].[member_id]=o.[member_id]  "
'-------------------------------

'-------------------------------
' Orders Open Event begin
' Orders Open Event end
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
    SetVar "DListOrders", ""
    Parse "OrdersNoRecords", False
    SetVar "OrdersNavigator", ""
    Parse "FormOrders", False
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
  iPage = GetParam("FormOrders_Page")
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
    fldField1_URLLink = "OrdersRecord.asp"
    fldField1_order_id = GetValue(rs, "o_order_id")
    flditem_id = GetValue(rs, "i_name")
    fldmember_id = GetValue(rs, "m_member_login")
    fldorder_id = GetValue(rs, "o_order_id")
    fldquantity = GetValue(rs, "o_quantity")
    fldField1= "Edit"
'-------------------------------
' Orders Show begin
'-------------------------------

'-------------------------------
' Orders Show Event begin
' Orders Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "Field1", ToHTML(fldField1)
      SetVar "Field1_URLLink", fldField1_URLLink
      SetVar "PrmField1_order_id", ToURL(fldField1_order_id)
      SetVar "order_id", ToHTML(fldorder_id)
      SetVar "item_id", ToHTML(flditem_id)
      SetVar "member_id", ToHTML(fldmember_id)
      SetVar "quantity", ToHTML(fldquantity)
    Parse "DListOrders", True

'-------------------------------
' Orders Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' Orders Navigation begin
'-------------------------------
  bEof = rs.eof
  if rs.eof and iPage = 1 then
	SetVar "OrdersNavigator", ""
  else
    if bEof then
      SetVar "OrdersNavigatorLastPage", "_"
    else
      SetVar "NextPage", (iPage + 1)
    end if
    if iPage = 1 then
      SetVar "OrdersNavigatorFirstPage", "_"
    else
      SetVar "PrevPage", (iPage - 1)
    end if
    SetVar "OrdersCurrentPage", iPage
    Parse "OrdersNavigator", False
  end if
'-------------------------------
' Orders Navigation end
'-------------------------------

'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "OrdersNoRecords", ""
  Parse "FormOrders", False

'-------------------------------
' Orders Close Event begin
' Orders Close Event end
'-------------------------------
End Sub
'===============================

%>