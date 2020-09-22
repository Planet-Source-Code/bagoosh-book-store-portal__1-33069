<%
'
'    Filename: MyInfo.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' MyInfo CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' MyInfo CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "MyInfo.asp"
sTemplateFileName = "MyInfo.html"
'===============================


'===============================
' MyInfo PageSecurity begin
CheckSecurity(1)
' MyInfo PageSecurity end
'===============================

'===============================
' MyInfo Open Event begin
' MyInfo Open Event end
'===============================

'===============================
' MyInfo OpenAnyPage Event begin
' MyInfo OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' MyInfo Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sFormErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Form"
    FormAction(sAction)
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
Form_Show
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

' MyInfo Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' MyInfo Close Event begin
' MyInfo Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub FormAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKmember_id : pPKmember_id = ""
  Dim fldmember_password : fldmember_password = ""
  Dim fldname : fldname = ""
  Dim fldlast_name : fldlast_name = ""
  Dim fldemail : fldemail = ""
  Dim fldaddress : fldaddress = ""
  Dim fldphone : fldphone = ""
  Dim fldnotes : fldnotes = ""
  Dim fldcard_type_id : fldcard_type_id = ""
  Dim fldcard_number : fldcard_number = ""
'-------------------------------

'-------------------------------
' Form Action begin
'-------------------------------
  sActionFileName = "ShoppingCart.asp"

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' Form BeforeCancel Event begin
' Form BeforeCancel Event end
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
    pPKmember_id = GetParam("PK_member_id")
    if IsEmpty(pPKmember_id) then exit sub
    sWhere = "member_id=" & ToSQL(pPKmember_id, "Number")
  end if
'-------------------------------


'-------------------------------
' Load all form fields into variables
'-------------------------------
  fldUserID = Session("UserID")
  fldmember_password = GetParam("member_password")
  fldname = GetParam("name")
  fldlast_name = GetParam("last_name")
  fldemail = GetParam("email")
  fldaddress = GetParam("address")
  fldphone = GetParam("phone")
  fldnotes = GetParam("notes")
  fldcard_type_id = GetParam("card_type_id")
  fldcard_number = GetParam("card_number")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldmember_password) then
      sFormErr = sFormErr & "The value in field Password* is required.<br>"
    end if
    if IsEmpty(fldname) then
      sFormErr = sFormErr & "The value in field First Name* is required.<br>"
    end if
    if IsEmpty(fldlast_name) then
      sFormErr = sFormErr & "The value in field Last Name* is required.<br>"
    end if
    if IsEmpty(fldemail) then
      sFormErr = sFormErr & "The value in field Email* is required.<br>"
    end if
    if not isNumeric(fldcard_type_id) then
      sFormErr = sFormErr & "The value in field Credit Card Type is incorrect.<br>"
    end if
'-------------------------------
' Form Check Event begin
' Form Check Event end
'-------------------------------
    If len(sFormErr) > 0 then
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
' Form Update Event begin
' Form Update Event end
'-------------------------------
      sSQL = "update members set " & _
        "[member_password]=" & ToSQL(fldmember_password, "Text") & _
        ",[first_name]=" & ToSQL(fldname, "Text") & _
        ",[last_name]=" & ToSQL(fldlast_name, "Text") & _
        ",[email]=" & ToSQL(fldemail, "Text") & _
        ",[address]=" & ToSQL(fldaddress, "Text") & _
        ",[phone]=" & ToSQL(fldphone, "Text") & _
        ",[notes]=" & ToSQL(fldnotes, "Text") & _
        ",[card_type_id]=" & ToSQL(fldcard_type_id, "Number") & _
        ",[card_number]=" & ToSQL(fldcard_number, "Text")
      sSQL = sSQL & " where " & sWhere
  end select
'-------------------------------
'-------------------------------
' Form BeforeExecute Event begin
' Form BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sFormErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sFormErr = ProcessError
  on error goto 0
  if len(sFormErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' Form Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Form_Show()
'-------------------------------
' Form Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "MyInfo"
  Dim bPK : bPK = True
  Dim scard_type_idDisplayValue: scard_type_idDisplayValue = ""

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sFormErr = "" then
    SetVar "FormError", ""
  else
    fldmember_id = GetParam("member_id")
    fldmember_password = GetParam("member_password")
    fldname = GetParam("name")
    fldlast_name = GetParam("last_name")
    fldemail = GetParam("email")
    fldaddress = GetParam("address")
    fldphone = GetParam("phone")
    fldnotes = GetParam("notes")
    fldcard_type_id = GetParam("card_type_id")
    fldcard_number = GetParam("card_number")
    SetVar "sFormErr", sFormErr
    SetVar "FormTitle", sFormTitle
    Parse "FormError", False
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
' Form Open Event begin
' Form Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from members where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Form") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldmember_id = GetValue(rs, "member_id")
    fldmember_login = GetValue(rs, "member_login")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sFormErr = "" then
      fldmember_password = GetValue(rs, "member_password")
      fldname = GetValue(rs, "first_name")
      fldlast_name = GetValue(rs, "last_name")
      fldemail = GetValue(rs, "email")
      fldaddress = GetValue(rs, "address")
      fldphone = GetValue(rs, "phone")
      fldnotes = GetValue(rs, "notes")
      fldcard_type_id = GetValue(rs, "card_type_id")
      fldcard_number = GetValue(rs, "card_number")
    end if
    SetVar "FormDelete", ""
    SetVar "FormInsert", ""
    Parse "FormEdit", False
'-------------------------------
' Form ShowEdit Event begin
' Form ShowEdit Event end
'-------------------------------
  else
    if sFormErr = "" then
      fldmember_id = ToHTML(Session("UserID"))
    end if
    SetVar "FormEdit", ""
    SetVar "FormInsert", ""
'-------------------------------
' Form ShowInsert Event begin
' Form ShowInsert Event end
'-------------------------------
  end if
  Parse "FormCancel", false
'-------------------------------
' Form Show Event begin
' Form Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "member_id", ToHTML(fldmember_id)
      SetVar "member_login", ToHTML(fldmember_login)
      SetVar "member_password", ToHTML(fldmember_password)
      SetVar "name", ToHTML(fldname)
      SetVar "last_name", ToHTML(fldlast_name)
      SetVar "email", ToHTML(fldemail)
      SetVar "address", ToHTML(fldaddress)
      SetVar "phone", ToHTML(fldphone)
      SetVar "notes", ToHTML(fldnotes)
      SetVar "FormLBcard_type_id", ""
      SetVar "Selected", ""
      SetVar "ID", ""
      SetVar "Value", scard_type_idDisplayValue
      Parse "FormLBcard_type_id", True
      openrs rscard_type_id, "select card_type_id, name from card_types order by 2"
      while not rscard_type_id.EOF
        SetVar "ID", GetValue(rscard_type_id, 0) : SetVar "Value", GetValue(rscard_type_id, 1)
        if cstr(GetValue(rscard_type_id, 0)) = cstr(fldcard_type_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "FormLBcard_type_id", True
        rscard_type_id.MoveNext
      wend
      set rscard_type_id = nothing
    
      SetVar "card_number", ToHTML(fldcard_number)
  Parse "FormForm", False

'-------------------------------
' Form Close Event begin
' Form Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Form Show end
'-------------------------------
End Sub
'===============================
%>