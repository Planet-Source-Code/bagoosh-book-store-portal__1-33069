<%
'
'    Filename: ShoppingCart.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' ShoppingCart CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' ShoppingCart CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "ShoppingCart.asp"
sTemplateFileName = "ShoppingCart.html"
'===============================


'===============================
' ShoppingCart PageSecurity begin
CheckSecurity(1)
' ShoppingCart PageSecurity end
'===============================

'===============================
' ShoppingCart Open Event begin
' ShoppingCart Open Event end
'===============================

'===============================
' ShoppingCart OpenAnyPage Event begin
' ShoppingCart OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' ShoppingCart Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sMemberErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Member"
    MemberAction(sAction)
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
Member_Show
Items_Show
Total_Show
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

' ShoppingCart Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' ShoppingCart Close Event begin
' ShoppingCart Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
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
  Dim sFormTitle: sFormTitle = "Items"
  Dim HasParam : HasParam = false
  Dim bReq : bReq = true
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", ""
  SetVar "FormParams", ""

'-------------------------------
' Build WHERE statement
'-------------------------------
  pUserID = Session("UserID")
  if IsNumeric(pUserID) and not isEmpty(pUserID) then pUserID = ToSQL(pUserID, "Number") else pUserID = Empty
  if not isEmpty(pUserID) then
    HasParam = true
    sWhere = sWhere & "[member_id]=" & pUserID
  else
    bReq = false
  end if


  if HasParam then
    sWhere = " AND (" & sWhere & ")"
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "SELECT order_id, name, price, quantity, member_id, quantity*price as sub_total FROM items, orders WHERE orders.item_id=items.item_id"
  sOrder = " ORDER BY order_id"
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
' Process if form has all required parameter
'-------------------------------
  if not bReq then
    SetVar "DListItems", ""
    Parse "ItemsNoRecords", False
    Parse "FormItems", False
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
    SetVar "DListItems", ""
    Parse "ItemsNoRecords", False
    Parse "FormItems", False
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
    fldField1_URLLink = "ShoppingCartRecord.asp"
    fldField1_order_id = GetValue(rs, "order_id")
    flditem_id = GetValue(rs, "name")
    fldorder_id = GetValue(rs, "order_id")
    fldprice = GetValue(rs, "price")
    fldquantity = GetValue(rs, "quantity")
    fldsub_total = GetValue(rs, "sub_total")
    fldField1= "Details"
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
      SetVar "PrmField1_order_id", ToURL(fldField1_order_id)
      SetVar "order_id", ToHTML(fldorder_id)
      SetVar "item_id", ToHTML(flditem_id)
      SetVar "price", ToHTML(fldprice)
      SetVar "quantity", ToHTML(fldquantity)
      SetVar "sub_total", ToHTML(fldsub_total)
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
  Dim bReq : bReq = true
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0

  SetVar "TransitParams", ""
  SetVar "FormParams", ""

'-------------------------------
' Build WHERE statement
'-------------------------------
  pUserID = Session("UserID")
  if IsNumeric(pUserID) and not isEmpty(pUserID) then pUserID = ToSQL(pUserID, "Number") else pUserID = Empty
  if not isEmpty(pUserID) then
    HasParam = true
    sWhere = sWhere & "[member_id]=" & pUserID
  else
    bReq = false
  end if


  if HasParam then
    sWhere = " AND (" & sWhere & ")"
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "SELECT member_id, sum(quantity*price) as sub_total FROM items, orders WHERE orders.item_id=items.item_id"
  sOrder = " GROUP BY member_id"
'-------------------------------

'-------------------------------
' Total Open Event begin
' Total Open Event end
'-------------------------------

'-------------------------------
' Assemble full SQL statement
'-------------------------------
  sSQL = sSQL & sWhere & sOrder
'-------------------------------

SetVar "FormTitle", sFormTitle

'-------------------------------
' Process if form has all required parameter
'-------------------------------
  if not bReq then
    SetVar "DListTotal", ""
    Parse "TotalNoRecords", False
    Parse "FormTotal", False
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
    fldsub_total = GetValue(rs, "sub_total")
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
    
      SetVar "sub_total", ToHTML(fldsub_total)
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

'===============================
' Action of the Record Form
'-------------------------------
Sub MemberAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sParams : sParams = "?"
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKmember_id : pPKmember_id = ""
'-------------------------------

'-------------------------------
' Member Action begin
'-------------------------------
  sActionFileName = "AdminMenu.asp"
  sParams = sParams & "UserID=" & ToURL(GetParam("Trn_UserID"))

'-------------------------------
' Load all form fields into variables
'-------------------------------
  fldUserID = Session("UserID")
'-------------------------------
' Member BeforeExecute Event begin
' Member BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sMemberErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sMemberErr = ProcessError
  on error goto 0
  if len(sMemberErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName & sParams
'-------------------------------
' Member Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Member_Show()
'-------------------------------
' Member Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "User Information"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sMemberErr = "" then
    SetVar "MemberError", ""
  else
    fldmember_id = GetParam("member_id")
    SetVar "sMemberErr", sMemberErr
    SetVar "FormTitle", sFormTitle
    Parse "MemberError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

'-------------------------------

'-------------------------------
' Build WHERE statement

  pmember_id = Session("UserID")
  if IsEmpty(pmember_id) then bPK = False
  
  sWhere = sWhere & "member_id=" & ToSQL(pmember_id, "Number")
  SetVar "PK_member_id", pmember_id
'-------------------------------
'-------------------------------
' Member Open Event begin
' Member Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from members where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Member") and not rs.eof)
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
    fldmember_login_URLLink = "MyInfo.asp"
    fldmember_login = GetValue(rs, "member_login")
    fldphone = GetValue(rs, "phone")
    SetVar "MemberDelete", ""
    SetVar "MemberUpdate", ""
    SetVar "MemberInsert", ""
'-------------------------------
' Member ShowEdit Event begin
' Member ShowEdit Event end
'-------------------------------
  else
    if sMemberErr = "" then
      fldmember_id = ToHTML(Session("UserID"))
    end if
    SetVar "MemberEdit", ""
    SetVar "MemberInsert", ""
'-------------------------------
' Member ShowInsert Event begin
' Member ShowInsert Event end
'-------------------------------
  end if
  SetVar "MemberCancel", ""
'-------------------------------
' Member Show Event begin
' Member Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "member_id", ToHTML(fldmember_id)
      SetVar "member_login", ToHTML(fldmember_login)
      SetVar "member_login_URLLink", fldmember_login_URLLink
      SetVar "name", ToHTML(fldname)
      SetVar "last_name", ToHTML(fldlast_name)
      SetVar "address", ToHTML(fldaddress)
      SetVar "email", ToHTML(fldemail)
      SetVar "phone", ToHTML(fldphone)
  Parse "FormMember", False

'-------------------------------
' Member Close Event begin
' Member Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Member Show end
'-------------------------------
End Sub
'===============================
%>