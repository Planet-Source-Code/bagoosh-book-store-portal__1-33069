<%
'
'    Filename: CardTypesRecord.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' CardTypesRecord CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' CardTypesRecord CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "CardTypesRecord.asp"
sTemplateFileName = "CardTypesRecord.html"
'===============================


'===============================
' CardTypesRecord PageSecurity begin
CheckSecurity(2)
' CardTypesRecord PageSecurity end
'===============================

'===============================
' CardTypesRecord Open Event begin
' CardTypesRecord Open Event end
'===============================

'===============================
' CardTypesRecord OpenAnyPage Event begin
' CardTypesRecord OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' CardTypesRecord Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sCardTypesErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "CardTypes"
    CardTypesAction(sAction)
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
CardTypes_Show
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

' CardTypesRecord Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' CardTypesRecord Close Event begin
' CardTypesRecord Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub CardTypesAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKcard_type_id : pPKcard_type_id = ""
  Dim fldname : fldname = ""
'-------------------------------

'-------------------------------
' CardTypes Action begin
'-------------------------------
  sActionFileName = "CardTypesGrid.asp"

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' CardTypes BeforeCancel Event begin
' CardTypes BeforeCancel Event end
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
    pPKcard_type_id = GetParam("PK_card_type_id")
    if IsEmpty(pPKcard_type_id) then exit sub
    sWhere = "card_type_id=" & ToSQL(pPKcard_type_id, "Number")
  end if
'-------------------------------


'-------------------------------
' Load all form fields into variables
'-------------------------------
  fldname = GetParam("name")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldname) then
      sCardTypesErr = sCardTypesErr & "The value in field Name is required.<br>"
    end if
'-------------------------------
' CardTypes Check Event begin
' CardTypes Check Event end
'-------------------------------
    If len(sCardTypesErr) > 0 then
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
' CardTypes Insert Event begin
' CardTypes Insert Event end
'-------------------------------
      sSQL = "insert into card_types (" & _
          "[name])" & _
          " values (" & _
          ToSQL(fldname, "Text") & _
          ")"
    case "update"
'-------------------------------
' CardTypes Update Event begin
' CardTypes Update Event end
'-------------------------------
      sSQL = "update card_types set " & _
        "[name]=" & ToSQL(fldname, "Text")
      sSQL = sSQL & " where " & sWhere
    case "delete"
'-------------------------------
' CardTypes Delete Event begin
' CardTypes Delete Event end
'-------------------------------
      sSQL = "delete from card_types where " & sWhere
  end select
'-------------------------------
'-------------------------------
' CardTypes BeforeExecute Event begin
' CardTypes BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sCardTypesErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sCardTypesErr = ProcessError
  on error goto 0
  if len(sCardTypesErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' CardTypes Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub CardTypes_Show()
'-------------------------------
' CardTypes Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Card Type"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sCardTypesErr = "" then
    fldcard_type_id = GetParam("card_type_id")
    pcard_type_id = GetParam("card_type_id")
    SetVar "CardTypesError", ""
  else
    fldcard_type_id = GetParam("card_type_id")
    fldname = GetParam("name")
    pcard_type_id = GetParam("PK_card_type_id")
    SetVar "sCardTypesErr", sCardTypesErr
    SetVar "FormTitle", sFormTitle
    Parse "CardTypesError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

'-------------------------------

'-------------------------------
' Build WHERE statement

  if IsEmpty(pcard_type_id) then bPK = False
  
  sWhere = sWhere & "card_type_id=" & ToSQL(pcard_type_id, "Number")
  SetVar "PK_card_type_id", pcard_type_id
'-------------------------------
'-------------------------------
' CardTypes Open Event begin
' CardTypes Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from card_types where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "CardTypes") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldcard_type_id = GetValue(rs, "card_type_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sCardTypesErr = "" then
      fldname = GetValue(rs, "name")
    end if
    SetVar "CardTypesInsert", ""
    Parse "CardTypesEdit", False
'-------------------------------
' CardTypes ShowEdit Event begin
' CardTypes ShowEdit Event end
'-------------------------------
  else
    if sCardTypesErr = "" then
      fldcard_type_id = ToHTML(GetParam("card_type_id"))
    end if
    SetVar "CardTypesEdit", ""
    Parse "CardTypesInsert", False
'-------------------------------
' CardTypes ShowInsert Event begin
' CardTypes ShowInsert Event end
'-------------------------------
  end if
  Parse "CardTypesCancel", false
'-------------------------------
' CardTypes Show Event begin
' CardTypes Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "card_type_id", ToHTML(fldcard_type_id)
      SetVar "name", ToHTML(fldname)
  Parse "FormCardTypes", False

'-------------------------------
' CardTypes Close Event begin
' CardTypes Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' CardTypes Show end
'-------------------------------
End Sub
'===============================
%>