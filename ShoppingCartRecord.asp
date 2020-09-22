<%
'
'    Filename: ShoppingCartRecord.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' ShoppingCartRecord CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' ShoppingCartRecord CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "ShoppingCartRecord.asp"
sTemplateFileName = "ShoppingCartRecord.html"
'===============================


'===============================
' ShoppingCartRecord PageSecurity begin
CheckSecurity(1)
' ShoppingCartRecord PageSecurity end
'===============================

'===============================
' ShoppingCartRecord Open Event begin
' ShoppingCartRecord Open Event end
'===============================

'===============================
' ShoppingCartRecord OpenAnyPage Event begin
' ShoppingCartRecord OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' ShoppingCartRecord Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sShoppingCartRecordErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "ShoppingCartRecord"
    ShoppingCartRecordAction(sAction)
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
ShoppingCartRecord_Show
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

' ShoppingCartRecord Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' ShoppingCartRecord Close Event begin
' ShoppingCartRecord Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub ShoppingCartRecordAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKorder_id : pPKorder_id = ""
  Dim fldmember_id : fldmember_id = ""
  Dim fldquantity : fldquantity = ""
'-------------------------------

'-------------------------------
' ShoppingCartRecord Action begin
'-------------------------------
  sActionFileName = "ShoppingCart.asp"

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' ShoppingCartRecord BeforeCancel Event begin
' ShoppingCartRecord BeforeCancel Event end
'-------------------------------
    cn.Close
    Set cn = Nothing
    response.redirect sActionFileName
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
  fldUserID = Session("UserID")
  fldmember_id = GetParam("member_id")
  fldquantity = GetParam("quantity")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldquantity) then
      sShoppingCartRecordErr = sShoppingCartRecordErr & "The value in field Quantity is required.<br>"
    end if
    if not isNumeric(fldmember_id) then
      sShoppingCartRecordErr = sShoppingCartRecordErr & "The value in field member_id is incorrect.<br>"
    end if
    if not isNumeric(fldquantity) then
      sShoppingCartRecordErr = sShoppingCartRecordErr & "The value in field Quantity is incorrect.<br>"
    end if
'-------------------------------
' ShoppingCartRecord Check Event begin
' ShoppingCartRecord Check Event end
'-------------------------------
    If len(sShoppingCartRecordErr) > 0 then
      exit sub
    end if
  end if
'-------------------------------


'-------------------------------
' Create SQL statement
'-------------------------------
  select case sAction
    case "update"
'-------------------------------
' ShoppingCartRecord Update Event begin
' ShoppingCartRecord Update Event end
'-------------------------------
      sSQL = "update orders set " & _
        "[member_id]=" & ToSQL(fldmember_id, "Number") & _
        ",[quantity]=" & ToSQL(fldquantity, "Number")
      sSQL = sSQL & " where " & sWhere
    case "delete"
'-------------------------------
' ShoppingCartRecord Delete Event begin
' ShoppingCartRecord Delete Event end
'-------------------------------
      sSQL = "delete from orders where " & sWhere
  end select
'-------------------------------
'-------------------------------
' ShoppingCartRecord BeforeExecute Event begin
' ShoppingCartRecord BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sShoppingCartRecordErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sShoppingCartRecordErr = ProcessError
  on error goto 0
  if len(sShoppingCartRecordErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' ShoppingCartRecord Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub ShoppingCartRecord_Show()
'-------------------------------
' ShoppingCartRecord Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "ShoppingCart"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sShoppingCartRecordErr = "" then
    porder_id = GetParam("order_id")
    SetVar "ShoppingCartRecordError", ""
  else
    fldorder_id = GetParam("order_id")
    fldmember_id = GetParam("member_id")
    fldquantity = GetParam("quantity")
    porder_id = GetParam("PK_order_id")
    SetVar "sShoppingCartRecordErr", sShoppingCartRecordErr
    SetVar "FormTitle", sFormTitle
    Parse "ShoppingCartRecordError", False
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
' ShoppingCartRecord Open Event begin
' ShoppingCartRecord Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from orders where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "ShoppingCartRecord") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    flditem_id = GetValue(rs, "item_id")
    fldmember_id = GetValue(rs, "member_id")
    fldorder_id = GetValue(rs, "order_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sShoppingCartRecordErr = "" then
      fldquantity = GetValue(rs, "quantity")
    end if
    SetVar "ShoppingCartRecordInsert", ""
    Parse "ShoppingCartRecordEdit", False
'-------------------------------
' ShoppingCartRecord ShowEdit Event begin
if cLng(fldmember_id) <> cLng(session("UserID")) then
CheckSecurity(2)
end if
' ShoppingCartRecord ShowEdit Event end
'-------------------------------
  else
    if sShoppingCartRecordErr = "" then
      fldmember_id = ToHTML(Session("UserID"))
    end if
    SetVar "ShoppingCartRecordEdit", ""
    SetVar "ShoppingCartRecordInsert", ""
'-------------------------------
' ShoppingCartRecord ShowInsert Event begin
' ShoppingCartRecord ShowInsert Event end
'-------------------------------
  end if
  Parse "ShoppingCartRecordCancel", false

'-------------------------------
' Set lookup fields
'-------------------------------
  flditem_id = DLookUp("items", "name", "item_id=" & ToSQL(flditem_id, "Number"))
'-------------------------------
' ShoppingCartRecord Show Event begin
' ShoppingCartRecord Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "order_id", ToHTML(fldorder_id)
      SetVar "member_id", ToHTML(fldmember_id)
      SetVar "item_id", ToHTML(flditem_id)
      SetVar "quantity", ToHTML(fldquantity)
  Parse "FormShoppingCartRecord", False

'-------------------------------
' ShoppingCartRecord Close Event begin
' ShoppingCartRecord Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' ShoppingCartRecord Show end
'-------------------------------
End Sub
'===============================
%>