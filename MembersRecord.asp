<%
'
'    Filename: MembersRecord.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' MembersRecord CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' MembersRecord CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "MembersRecord.asp"
sTemplateFileName = "MembersRecord.html"
'===============================


'===============================
' MembersRecord PageSecurity begin
CheckSecurity(2)
' MembersRecord PageSecurity end
'===============================

'===============================
' MembersRecord Open Event begin
' MembersRecord Open Event end
'===============================

'===============================
' MembersRecord OpenAnyPage Event begin
' MembersRecord OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' MembersRecord Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sMembersErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Members"
    MembersAction(sAction)
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
Members_Show
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

' MembersRecord Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' MembersRecord Close Event begin
' MembersRecord Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub MembersAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sParams : sParams = "?"
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKmember_id : pPKmember_id = ""
  Dim fldmember_login : fldmember_login = ""
  Dim fldmember_password : fldmember_password = ""
  Dim fldmember_level : fldmember_level = ""
  Dim fldname : fldname = ""
  Dim fldlast_name : fldlast_name = ""
  Dim fldemail : fldemail = ""
  Dim fldphone : fldphone = ""
  Dim fldaddress : fldaddress = ""
  Dim fldnotes : fldnotes = ""
  Dim fldcard_type_id : fldcard_type_id = ""
  Dim fldcard_number : fldcard_number = ""
'-------------------------------

'-------------------------------
' Members Action begin
'-------------------------------
  sActionFileName = "MembersGrid.asp"
  sParams = sParams & "member_login=" & ToURL(GetParam("Trn_member_login"))

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' Members BeforeCancel Event begin
' Members BeforeCancel Event end
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
  fldmember_level = GetParam("member_level")
  fldname = GetParam("name")
  fldlast_name = GetParam("last_name")
  fldemail = GetParam("email")
  fldphone = GetParam("phone")
  fldaddress = GetParam("address")
  fldnotes = GetParam("notes")
  fldcard_type_id = GetParam("card_type_id")
  fldcard_number = GetParam("card_number")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldmember_login) then
      sMembersErr = sMembersErr & "The value in field Login* is required.<br>"
    end if
    if IsEmpty(fldmember_password) then
      sMembersErr = sMembersErr & "The value in field Password* is required.<br>"
    end if
    if IsEmpty(fldmember_level) then
      sMembersErr = sMembersErr & "The value in field Level* is required.<br>"
    end if
    if IsEmpty(fldname) then
      sMembersErr = sMembersErr & "The value in field First Name* is required.<br>"
    end if
    if IsEmpty(fldlast_name) then
      sMembersErr = sMembersErr & "The value in field Last Name* is required.<br>"
    end if
    if IsEmpty(fldemail) then
      sMembersErr = sMembersErr & "The value in field Email* is required.<br>"
    end if
    if not isNumeric(fldmember_level) then
      sMembersErr = sMembersErr & "The value in field Level* is incorrect.<br>"
    end if
    if not isNumeric(fldcard_type_id) then
      sMembersErr = sMembersErr & "The value in field Credit Card Type is incorrect.<br>"
    end if
    if not IsEmpty(fldmember_login) then
      iCount = 0
      if sAction = "insert" then
        iCount = Clng(DLookUp("members", "count(*)", "member_login=" & toSQL(fldmember_login, "Text")))
      elseif sAction = "update" then
        iCount = Clng(DLookUp("members", "count(*)", "member_login=" & toSQL(fldmember_login, "Text") & " and not(" & sWhere & ")"))
      end if
      if iCount > 0 then
        sMembersErr = sMembersErr & "The value in field Login* is already in database.<br>"
      end if
    end if
'-------------------------------
' Members Check Event begin
' Members Check Event end
'-------------------------------
    If len(sMembersErr) > 0 then
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
' Members Insert Event begin
' Members Insert Event end
'-------------------------------
      sSQL = "insert into members (" & _
          "[member_login]," & _
          "[member_password]," & _
          "[member_level]," & _
          "[first_name]," & _
          "[last_name]," & _
          "[email]," & _
          "[phone]," & _
          "[address]," & _
          "[notes]," & _
          "[card_type_id]," & _
          "[card_number])" & _
          " values (" & _
          ToSQL(fldmember_login, "Text") & "," & _
          ToSQL(fldmember_password, "Text") & "," & _
          ToSQL(fldmember_level, "Number") & "," & _
          ToSQL(fldname, "Text") & "," & _
          ToSQL(fldlast_name, "Text") & "," & _
          ToSQL(fldemail, "Text") & "," & _
          ToSQL(fldphone, "Text") & "," & _
          ToSQL(fldaddress, "Text") & "," & _
          ToSQL(fldnotes, "Text") & "," & _
          ToSQL(fldcard_type_id, "Number") & "," & _
          ToSQL(fldcard_number, "Text") & _
          ")"
    case "update"
'-------------------------------
' Members Update Event begin
' Members Update Event end
'-------------------------------
      sSQL = "update members set " & _
        "[member_login]=" & ToSQL(fldmember_login, "Text") & _
        ",[member_password]=" & ToSQL(fldmember_password, "Text") & _
        ",[member_level]=" & ToSQL(fldmember_level, "Number") & _
        ",[first_name]=" & ToSQL(fldname, "Text") & _
        ",[last_name]=" & ToSQL(fldlast_name, "Text") & _
        ",[email]=" & ToSQL(fldemail, "Text") & _
        ",[phone]=" & ToSQL(fldphone, "Text") & _
        ",[address]=" & ToSQL(fldaddress, "Text") & _
        ",[notes]=" & ToSQL(fldnotes, "Text") & _
        ",[card_type_id]=" & ToSQL(fldcard_type_id, "Number") & _
        ",[card_number]=" & ToSQL(fldcard_number, "Text")
      sSQL = sSQL & " where " & sWhere
    case "delete"
'-------------------------------
' Members Delete Event begin
' Members Delete Event end
'-------------------------------
      sSQL = "delete from members where " & sWhere
  end select
'-------------------------------
'-------------------------------
' Members BeforeExecute Event begin
' Members BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sMembersErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sMembersErr = ProcessError
  on error goto 0
  if len(sMembersErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName & sParams
'-------------------------------
' Members Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Members_Show()
'-------------------------------
' Members Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Members"
  Dim bPK : bPK = True
  Dim scard_type_idDisplayValue: scard_type_idDisplayValue = ""

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sMembersErr = "" then
    fldmember_login = GetParam("member_login")
    fldmember_id = GetParam("member_id")
    SetVar "Trn_member_login", GetParam("member_login")
    pmember_id = GetParam("member_id")
    SetVar "MembersError", ""
  else
    fldmember_id = GetParam("member_id")
    fldmember_login = GetParam("member_login")
    fldmember_password = GetParam("member_password")
    fldmember_level = GetParam("member_level")
    fldname = GetParam("name")
    fldlast_name = GetParam("last_name")
    fldemail = GetParam("email")
    fldphone = GetParam("phone")
    fldaddress = GetParam("address")
    fldnotes = GetParam("notes")
    fldcard_type_id = GetParam("card_type_id")
    fldcard_number = GetParam("card_number")
    SetVar "Trn_member_login", GetParam("Trn_member_login")
    pmember_id = GetParam("PK_member_id")
    SetVar "sMembersErr", sMembersErr
    SetVar "FormTitle", sFormTitle
    Parse "MembersError", False
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
' Members Open Event begin
' Members Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from members where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Members") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldmember_id = GetValue(rs, "member_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sMembersErr = "" then
      fldmember_login = GetValue(rs, "member_login")
      fldmember_password = GetValue(rs, "member_password")
      fldmember_level = GetValue(rs, "member_level")
      fldname = GetValue(rs, "first_name")
      fldlast_name = GetValue(rs, "last_name")
      fldemail = GetValue(rs, "email")
      fldphone = GetValue(rs, "phone")
      fldaddress = GetValue(rs, "address")
      fldnotes = GetValue(rs, "notes")
      fldcard_type_id = GetValue(rs, "card_type_id")
      fldcard_number = GetValue(rs, "card_number")
    end if
    SetVar "MembersInsert", ""
    Parse "MembersEdit", False
'-------------------------------
' Members ShowEdit Event begin
' Members ShowEdit Event end
'-------------------------------
  else
    if sMembersErr = "" then
      fldmember_id = ToHTML(GetParam("member_id"))
      fldmember_login = ToHTML(GetParam("member_login"))
    end if
    SetVar "MembersEdit", ""
    Parse "MembersInsert", False
'-------------------------------
' Members ShowInsert Event begin
' Members ShowInsert Event end
'-------------------------------
  end if
  Parse "MembersCancel", false
'-------------------------------
' Members Show Event begin
' Members Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "member_id", ToHTML(fldmember_id)
      SetVar "member_login", ToHTML(fldmember_login)
      SetVar "member_password", ToHTML(fldmember_password)
      SetVar "MembersLBmember_level", ""
      LOV = Split("1;Member;2;Administrator", ";")
      if (ubound(LOV) mod 2) = 1 then
        for i = 0 to ubound(LOV) step 2
          SetVar "ID", LOV(i) : SetVar "Value", LOV(i+1)
          if cstr(LOV(i)) = cstr(fldmember_level) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
          Parse "MembersLBmember_level", True
        next
      end if
    
      SetVar "name", ToHTML(fldname)
      SetVar "last_name", ToHTML(fldlast_name)
      SetVar "email", ToHTML(fldemail)
      SetVar "phone", ToHTML(fldphone)
      SetVar "address", ToHTML(fldaddress)
      SetVar "notes", ToHTML(fldnotes)
      SetVar "MembersLBcard_type_id", ""
      SetVar "Selected", ""
      SetVar "ID", ""
      SetVar "Value", scard_type_idDisplayValue
      Parse "MembersLBcard_type_id", True
      openrs rscard_type_id, "select card_type_id, name from card_types order by 2"
      while not rscard_type_id.EOF
        SetVar "ID", GetValue(rscard_type_id, 0) : SetVar "Value", GetValue(rscard_type_id, 1)
        if cstr(GetValue(rscard_type_id, 0)) = cstr(fldcard_type_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "MembersLBcard_type_id", True
        rscard_type_id.MoveNext
      wend
      set rscard_type_id = nothing
    
      SetVar "card_number", ToHTML(fldcard_number)
  Parse "FormMembers", False

'-------------------------------
' Members Close Event begin
' Members Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Members Show end
'-------------------------------
End Sub
'===============================
%>