<%
'
'    Filename: MembersInfo.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' MembersInfo CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' MembersInfo CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "MembersInfo.asp"
sTemplateFileName = "MembersInfo.html"
'===============================


'===============================
' MembersInfo PageSecurity begin
CheckSecurity(2)
' MembersInfo PageSecurity end
'===============================

'===============================
' MembersInfo Open Event begin
' MembersInfo Open Event end
'===============================

'===============================
' MembersInfo OpenAnyPage Event begin
' MembersInfo OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' MembersInfo Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sRecordErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Record"
    RecordAction(sAction)
end select
'===============================

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
Record_Show
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

' MembersInfo Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' MembersInfo Close Event begin
' MembersInfo Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub RecordAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKmember_id : pPKmember_id = ""
'-------------------------------

'-------------------------------
' Record Action begin
'-------------------------------
  sActionFileName = "AdminMenu.asp"

'-------------------------------
' Load all form fields into variables
'-------------------------------
'-------------------------------
' Record BeforeExecute Event begin
' Record BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sRecordErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sRecordErr = ProcessError
  on error goto 0
  if len(sRecordErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' Record Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Record_Show()
'-------------------------------
' Record Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Member Info"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sRecordErr = "" then
    pmember_id = GetParam("member_id")
    SetVar "RecordError", ""
  else
    fldmember_id = GetParam("member_id")
    pmember_id = GetParam("PK_member_id")
    SetVar "sRecordErr", sRecordErr
    SetVar "FormTitle", sFormTitle
    Parse "RecordError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

'-------------------------------

'-------------------------------
' Build WHERE statement

  if IsEmpty(pmember_id) then bPK = False
  
  sWhere = sWhere & "member_id=" & ToSQL(pmember_id, "Number")
  SetVar "PK_member_id", pmember_id
'-------------------------------
'-------------------------------
' Record Open Event begin
' Record Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from members where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Record") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldaddress = GetValue(rs, "address")
    fldemail = GetValue(rs, "email")
    fldname = GetValue(rs, "first_name")
    fldlast_name = GetValue(rs, "last_name")
    fldmember_id = GetValue(rs, "member_id")
    fldmember_level = GetValue(rs, "member_level")
    fldmember_login_URLLink = "MembersRecord.asp"
    fldmember_login_member_id = GetValue(rs, "member_id")
    fldmember_login = GetValue(rs, "member_login")
    fldnotes = GetValue(rs, "notes")
    fldphone = GetValue(rs, "phone")
    SetVar "RecordDelete", ""
    SetVar "RecordUpdate", ""
    SetVar "RecordInsert", ""
'-------------------------------
' Record ShowEdit Event begin
' Record ShowEdit Event end
'-------------------------------
  else
    SetVar "RecordEdit", ""
    SetVar "RecordInsert", ""
'-------------------------------
' Record ShowInsert Event begin
' Record ShowInsert Event end
'-------------------------------
  end if
  SetVar "RecordCancel", ""
'-------------------------------
' Record Show Event begin
' Record Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "member_id", ToHTML(fldmember_id)
      SetVar "member_login", ToHTML(fldmember_login)
      SetVar "member_login_URLLink", fldmember_login_URLLink
      SetVar "Prmmember_login_member_id", ToURL(fldmember_login_member_id)
      SetVar "member_level", ToHTML(fldmember_level)
      SetVar "name", ToHTML(fldname)
      SetVar "last_name", ToHTML(fldlast_name)
      SetVar "email", ToHTML(fldemail)
      SetVar "phone", ToHTML(fldphone)
      SetVar "address", ToHTML(fldaddress)
      SetVar "notes", ToHTML(fldnotes)
  Parse "FormRecord", False

'-------------------------------
' Record Close Event begin
' Record Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Record Show end
'-------------------------------
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
  Dim sFormTitle: sFormTitle = "Shopping Cart"
  Dim HasParam : HasParam = false
  Dim bReq : bReq = true
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", "member_id=" & ToURL(GetParam("member_id")) & "&"
  SetVar "FormParams", "member_id=" & ToURL(GetParam("member_id")) & "&"

'-------------------------------
' Build WHERE statement
'-------------------------------
  pmember_id = GetParam("member_id")
  if IsNumeric(pmember_id) and not isEmpty(pmember_id) then pmember_id = ToSQL(pmember_id, "Number") else pmember_id = Empty
  if not isEmpty(pmember_id) then
    HasParam = true
    sWhere = sWhere & "o.[member_id]=" & pmember_id
  else
    bReq = false
  end if


  if HasParam then
    sWhere = " AND (" & sWhere & ")"
  end if
  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
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
    if iSort = 1 then sOrder = " order by o.[order_id]" & sDirection
    if iSort = 2 then sOrder = " order by i.[name]" & sDirection
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
    "[i].[name] as i_name " & _
    " from [orders] o, [items] i" & _
    " where [i].[item_id]=o.[item_id]  "
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
' Process the parameters for sorting
'-------------------------------
  SetVar "SortParams", sSortParams
'-------------------------------

'-------------------------------
' Process if form has all required parameter
'-------------------------------
  if not bReq then
    SetVar "DListOrders", ""
    Parse "OrdersNoRecords", False
    Parse "FormOrders", False
    exit sub
  end if
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
' Display grid based on recordset
'-------------------------------
  while not rs.EOF  and iCounter < iRecordsPerPage
'-------------------------------
' Create field variables based on database fields
'-------------------------------
    flditem_id = GetValue(rs, "i_name")
    fldorder_id_URLLink = "OrdersRecord.asp"
    fldorder_id_order_id = GetValue(rs, "o_order_id")
    fldorder_id = GetValue(rs, "o_order_id")
    fldquantity = GetValue(rs, "o_quantity")
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
    
      SetVar "order_id", ToHTML(fldorder_id)
      SetVar "order_id_URLLink", fldorder_id_URLLink
      SetVar "Prmorder_id_order_id", ToURL(fldorder_id_order_id)
      SetVar "item_id", ToHTML(flditem_id)
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