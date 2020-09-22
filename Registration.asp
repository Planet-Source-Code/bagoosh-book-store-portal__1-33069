<%
'
'    Filename: Registration.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' Registration CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' Registration CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "Registration.asp"
sTemplateFileName = "Registration.html"
'===============================


'===============================
' Registration PageSecurity begin
' Registration PageSecurity end
'===============================

'===============================
' Registration Open Event begin
' Registration Open Event end
'===============================

'===============================
' Registration OpenAnyPage Event begin
' Registration OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' Registration Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sRegErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Reg"
    RegAction(sAction)
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
Reg_Show
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

' Registration Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' Registration Close Event begin
' Registration Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub RegAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKmember_id : pPKmember_id = ""
  Dim fldmember_login : fldmember_login = ""
  Dim fldmember_password : fldmember_password = ""
  Dim fldmember_password2 : fldmember_password2 = ""
  Dim fldfirst_name : fldfirst_name = ""
  Dim fldlast_name : fldlast_name = ""
  Dim fldemail : fldemail = ""
  Dim fldaddress : fldaddress = ""
  Dim fldphone : fldphone = ""
  Dim fldcard_type_id : fldcard_type_id = ""
  Dim fldcard_number : fldcard_number = ""
'-------------------------------

'-------------------------------
' Reg Action begin
'-------------------------------
  sActionFileName = "Default.asp"

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' Reg BeforeCancel Event begin
' Reg BeforeCancel Event end
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
  fldmember_login = GetParam("member_login")
  fldmember_password = GetParam("member_password")
  fldmember_password2 = GetParam("member_password2")
  fldfirst_name = GetParam("first_name")
  fldlast_name = GetParam("last_name")
  fldemail = GetParam("email")
  fldaddress = GetParam("address")
  fldphone = GetParam("phone")
  fldcard_type_id = GetParam("card_type_id")
  fldcard_number = GetParam("card_number")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldmember_login) then
      sRegErr = sRegErr & "The value in field Login* is required.<br>"
    end if
    if IsEmpty(fldmember_password) then
      sRegErr = sRegErr & "The value in field Password* is required.<br>"
    end if
    if IsEmpty(fldmember_password2) then
      sRegErr = sRegErr & "The value in field Confirm Password* is required.<br>"
    end if
    if IsEmpty(fldfirst_name) then
      sRegErr = sRegErr & "The value in field First Name* is required.<br>"
    end if
    if IsEmpty(fldlast_name) then
      sRegErr = sRegErr & "The value in field Last Name* is required.<br>"
    end if
    if IsEmpty(fldemail) then
      sRegErr = sRegErr & "The value in field Email* is required.<br>"
    end if
    if not isNumeric(fldcard_type_id) then
      sRegErr = sRegErr & "The value in field Credit Card Type is incorrect.<br>"
    end if
    if not IsEmpty(fldmember_login) then
      iCount = 0
      if sAction = "insert" then
        iCount = Clng(DLookUp("members", "count(*)", "member_login=" & toSQL(fldmember_login, "Text")))
      elseif sAction = "update" then
        iCount = Clng(DLookUp("members", "count(*)", "member_login=" & toSQL(fldmember_login, "Text") & " and not(" & sWhere & ")"))
      end if
      if iCount > 0 then
        sRegErr = sRegErr & "The value in field Login* is already in database.<br>"
      end if
    end if
'-------------------------------
' Reg Check Event begin
if getParam("member_password") <> getParam("member_password2") then
sRegErr = sRegErr & chr(13) & "Password and Confirm Password fields don't match"
end if
' Reg Check Event end
'-------------------------------
    If len(sRegErr) > 0 then
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
' Reg Insert Event begin
' Reg Insert Event end
'-------------------------------
      sSQL = "insert into members (" & _
          "[member_login]," & _
          "[member_password]," & _
          "[first_name]," & _
          "[last_name]," & _
          "[email]," & _
          "[address]," & _
          "[phone]," & _
          "[card_type_id]," & _
          "[card_number])" & _
          " values (" & _
          ToSQL(fldmember_login, "Text") & "," & _
          ToSQL(fldmember_password, "Text") & "," & _
          ToSQL(fldfirst_name, "Text") & "," & _
          ToSQL(fldlast_name, "Text") & "," & _
          ToSQL(fldemail, "Text") & "," & _
          ToSQL(fldaddress, "Text") & "," & _
          ToSQL(fldphone, "Text") & "," & _
          ToSQL(fldcard_type_id, "Number") & "," & _
          ToSQL(fldcard_number, "Text") & _
          ")"
    case "update"
'-------------------------------
' Reg Update Event begin
' Reg Update Event end
'-------------------------------
      sSQL = "update members set " & _
        "[member_login]=" & ToSQL(fldmember_login, "Text") & _
        ",[member_password]=" & ToSQL(fldmember_password, "Text") & _
        ",[first_name]=" & ToSQL(fldfirst_name, "Text") & _
        ",[last_name]=" & ToSQL(fldlast_name, "Text") & _
        ",[email]=" & ToSQL(fldemail, "Text") & _
        ",[address]=" & ToSQL(fldaddress, "Text") & _
        ",[phone]=" & ToSQL(fldphone, "Text") & _
        ",[card_type_id]=" & ToSQL(fldcard_type_id, "Number") & _
        ",[card_number]=" & ToSQL(fldcard_number, "Text")
      sSQL = sSQL & " where " & sWhere
  end select
'-------------------------------
'-------------------------------
' Reg BeforeExecute Event begin
' Reg BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sRegErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sRegErr = ProcessError
  on error goto 0
  if len(sRegErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' Reg Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Reg_Show()
'-------------------------------
' Reg Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Registration"
  Dim bPK : bPK = True
  Dim scard_type_idDisplayValue: scard_type_idDisplayValue = ""

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sRegErr = "" then
    SetVar "RegError", ""
  else
    fldmember_id = GetParam("member_id")
    fldmember_login = GetParam("member_login")
    fldmember_password = GetParam("member_password")
    fldfirst_name = GetParam("first_name")
    fldlast_name = GetParam("last_name")
    fldemail = GetParam("email")
    fldaddress = GetParam("address")
    fldphone = GetParam("phone")
    fldcard_type_id = GetParam("card_type_id")
    fldcard_number = GetParam("card_number")
    SetVar "sRegErr", sRegErr
    SetVar "FormTitle", sFormTitle
    Parse "RegError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

  fldmember_password2 = GetParam("member_password2")
'-------------------------------

'-------------------------------
' Build WHERE statement

  pmember_id = Session("UserID")
  if IsEmpty(pmember_id) then bPK = False
  
  sWhere = sWhere & "member_id=" & ToSQL(pmember_id, "Number")
  SetVar "PK_member_id", pmember_id
'-------------------------------
'-------------------------------
' Reg Open Event begin
' Reg Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from members where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Reg") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldmember_id = GetValue(rs, "member_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sRegErr = "" then
      fldmember_login = GetValue(rs, "member_login")
      fldmember_password = GetValue(rs, "member_password")
      fldfirst_name = GetValue(rs, "first_name")
      fldlast_name = GetValue(rs, "last_name")
      fldemail = GetValue(rs, "email")
      fldaddress = GetValue(rs, "address")
      fldphone = GetValue(rs, "phone")
      fldcard_type_id = GetValue(rs, "card_type_id")
      fldcard_number = GetValue(rs, "card_number")
    end if
    SetVar "RegDelete", ""
    SetVar "RegInsert", ""
    Parse "RegEdit", False
'-------------------------------
' Reg ShowEdit Event begin
' Reg ShowEdit Event end
'-------------------------------
  else
    if sRegErr = "" then
      fldmember_id = ToHTML(Session("UserID"))
    end if
    SetVar "RegEdit", ""
    Parse "RegInsert", False
'-------------------------------
' Reg ShowInsert Event begin
' Reg ShowInsert Event end
'-------------------------------
  end if
  Parse "RegCancel", false
'-------------------------------
' Reg Show Event begin
' Reg Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "member_id", ToHTML(fldmember_id)
      SetVar "member_login", ToHTML(fldmember_login)
      SetVar "member_password", ToHTML(fldmember_password)
      SetVar "member_password2", ToHTML(fldmember_password2)
      SetVar "first_name", ToHTML(fldfirst_name)
      SetVar "last_name", ToHTML(fldlast_name)
      SetVar "email", ToHTML(fldemail)
      SetVar "address", ToHTML(fldaddress)
      SetVar "phone", ToHTML(fldphone)
      SetVar "RegLBcard_type_id", ""
      SetVar "Selected", ""
      SetVar "ID", ""
      SetVar "Value", scard_type_idDisplayValue
      Parse "RegLBcard_type_id", True
      openrs rscard_type_id, "select card_type_id, name from card_types order by 2"
      while not rscard_type_id.EOF
        SetVar "ID", GetValue(rscard_type_id, 0) : SetVar "Value", GetValue(rscard_type_id, 1)
        if cstr(GetValue(rscard_type_id, 0)) = cstr(fldcard_type_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "RegLBcard_type_id", True
        rscard_type_id.MoveNext
      wend
      set rscard_type_id = nothing
    
      SetVar "card_number", ToHTML(fldcard_number)
  Parse "FormReg", False

'-------------------------------
' Reg Close Event begin
' Reg Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Reg Show end
'-------------------------------
End Sub
'===============================
%>