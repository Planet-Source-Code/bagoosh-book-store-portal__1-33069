<%
'
'    Filename: OrdersRecord.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' OrdersRecord CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' OrdersRecord CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "OrdersRecord.asp"
sTemplateFileName = "OrdersRecord.html"
'===============================


'===============================
' OrdersRecord PageSecurity begin
CheckSecurity(2)
' OrdersRecord PageSecurity end
'===============================

'===============================
' OrdersRecord Open Event begin
' OrdersRecord Open Event end
'===============================

'===============================
' OrdersRecord OpenAnyPage Event begin
' OrdersRecord OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' OrdersRecord Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sOrdersErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Orders"
    OrdersAction(sAction)
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

' OrdersRecord Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' OrdersRecord Close Event begin
' OrdersRecord Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub OrdersAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sParams : sParams = "?"
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKorder_id : pPKorder_id = ""
  Dim fldmember_id : fldmember_id = ""
  Dim flditem_id : flditem_id = ""
  Dim fldquantity : fldquantity = ""
'-------------------------------

'-------------------------------
' Orders Action begin
'-------------------------------
  sActionFileName = "OrdersGrid.asp"
  sParams = sParams & "item_id=" & ToURL(GetParam("Trn_item_id")) & "&"
  sParams = sParams & "member_id=" & ToURL(GetParam("Trn_member_id"))

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' Orders BeforeCancel Event begin
' Orders BeforeCancel Event end
'-------------------------------
    cn.Close
    Set cn = Nothing
    response.redirect sActionFileName & sParams
  end if
'-------------------------------

'-------------------------------
' Build WHERE statement
'-------------------------------
  if sAction = "update" or sAction = "delete" then
    pPKorder_id = GetParam("PK_order_id")
    if IsEmpty(pPKorder_id) then exit sub
    sWhere = "order_id=" & ToSQL(pPKorder_id, "Number")
  end if
'-------------------------------


'-------------------------------
' Load all form fields into variables
'-------------------------------
  fldmember_id = GetParam("member_id")
  flditem_id = GetParam("item_id")
  fldquantity = GetParam("quantity")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldmember_id) then
      sOrdersErr = sOrdersErr & "The value in field Member is required.<br>"
    end if
    if IsEmpty(flditem_id) then
      sOrdersErr = sOrdersErr & "The value in field Item is required.<br>"
    end if
    if IsEmpty(fldquantity) then
      sOrdersErr = sOrdersErr & "The value in field Quantity is required.<br>"
    end if
    if not isNumeric(fldmember_id) then
      sOrdersErr = sOrdersErr & "The value in field Member is incorrect.<br>"
    end if
    if not isNumeric(flditem_id) then
      sOrdersErr = sOrdersErr & "The value in field Item is incorrect.<br>"
    end if
    if not isNumeric(fldquantity) then
      sOrdersErr = sOrdersErr & "The value in field Quantity is incorrect.<br>"
    end if
'-------------------------------
' Orders Check Event begin
' Orders Check Event end
'-------------------------------
    If len(sOrdersErr) > 0 then
      exit sub
    end if
  end if
'-------------------------------


'-------------------------------
' Create SQL statement
'-------------------------------
  select case sAction
    case "insert"
'-------------------------------
' Orders Insert Event begin
' Orders Insert Event end
'-------------------------------
      sSQL = "insert into orders (" & _
          "[member_id]," & _
          "[item_id]," & _
          "[quantity])" & _
          " values (" & _
          ToSQL(fldmember_id, "Number") & "," & _
          ToSQL(flditem_id, "Number") & "," & _
          ToSQL(fldquantity, "Number") & _
          ")"
    case "update"
'-------------------------------
' Orders Update Event begin
' Orders Update Event end
'-------------------------------
      sSQL = "update orders set " & _
        "[member_id]=" & ToSQL(fldmember_id, "Number") & _
        ",[item_id]=" & ToSQL(flditem_id, "Number") & _
        ",[quantity]=" & ToSQL(fldquantity, "Number")
      sSQL = sSQL & " where " & sWhere
    case "delete"
'-------------------------------
' Orders Delete Event begin
' Orders Delete Event end
'-------------------------------
      sSQL = "delete from orders where " & sWhere
  end select
'-------------------------------
'-------------------------------
' Orders BeforeExecute Event begin
' Orders BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sOrdersErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sOrdersErr = ProcessError
  on error goto 0
  if len(sOrdersErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName & sParams
'-------------------------------
' Orders Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Orders_Show()
'-------------------------------
' Orders Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Orders"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sOrdersErr = "" then
    flditem_id = GetParam("item_id")
    fldmember_id = GetParam("member_id")
    fldorder_id = GetParam("order_id")
    SetVar "Trn_item_id", GetParam("item_id")
    SetVar "Trn_member_id", GetParam("member_id")
    porder_id = GetParam("order_id")
    SetVar "OrdersError", ""
  else
    fldmember_id = GetParam("member_id")
    flditem_id = GetParam("item_id")
    fldquantity = GetParam("quantity")
    SetVar "Trn_item_id", GetParam("Trn_item_id")
    SetVar "Trn_member_id", GetParam("Trn_member_id")
    porder_id = GetParam("PK_order_id")
    SetVar "sOrdersErr", sOrdersErr
    SetVar "FormTitle", sFormTitle
    Parse "OrdersError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

'-------------------------------

'-------------------------------
' Build WHERE statement

  if IsEmpty(porder_id) then bPK = False
  
  sWhere = sWhere & "order_id=" & ToSQL(porder_id, "Number")
  SetVar "PK_order_id", porder_id
'-------------------------------
'-------------------------------
' Orders Open Event begin
' Orders Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from orders where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Orders") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldorder_id = GetValue(rs, "order_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sOrdersErr = "" then
      fldmember_id = GetValue(rs, "member_id")
      flditem_id = GetValue(rs, "item_id")
      fldquantity = GetValue(rs, "quantity")
    end if
    SetVar "OrdersInsert", ""
    Parse "OrdersEdit", False
'-------------------------------
' Orders ShowEdit Event begin
' Orders ShowEdit Event end
'-------------------------------
  else
    if sOrdersErr = "" then
      fldorder_id = ToHTML(GetParam("order_id"))
      fldmember_id = ToHTML(GetParam("member_id"))
      flditem_id = ToHTML(GetParam("item_id"))
    end if
    SetVar "OrdersEdit", ""
    Parse "OrdersInsert", False
'-------------------------------
' Orders ShowInsert Event begin
' Orders ShowInsert Event end
'-------------------------------
  end if
  Parse "OrdersCancel", false
'-------------------------------
' Orders Show Event begin
' Orders Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "order_id", ToHTML(fldorder_id)
      SetVar "OrdersLBmember_id", ""
      openrs rsmember_id, "select member_id, member_login from members order by 2"
      while not rsmember_id.EOF
        SetVar "ID", GetValue(rsmember_id, 0) : SetVar "Value", GetValue(rsmember_id, 1)
        if cstr(GetValue(rsmember_id, 0)) = cstr(fldmember_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "OrdersLBmember_id", True
        rsmember_id.MoveNext
      wend
      set rsmember_id = nothing
    
      SetVar "OrdersLBitem_id", ""
      openrs rsitem_id, "select item_id, name from items order by 2"
      while not rsitem_id.EOF
        SetVar "ID", GetValue(rsitem_id, 0) : SetVar "Value", GetValue(rsitem_id, 1)
        if cstr(GetValue(rsitem_id, 0)) = cstr(flditem_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "OrdersLBitem_id", True
        rsitem_id.MoveNext
      wend
      set rsitem_id = nothing
    
      SetVar "quantity", ToHTML(fldquantity)
  Parse "FormOrders", False

'-------------------------------
' Orders Close Event begin
' Orders Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Orders Show end
'-------------------------------
End Sub
'===============================
%>