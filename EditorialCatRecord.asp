<%
'
'    Filename: EditorialCatRecord.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' EditorialCatRecord CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' EditorialCatRecord CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "EditorialCatRecord.asp"
sTemplateFileName = "EditorialCatRecord.html"
'===============================


'===============================
' EditorialCatRecord PageSecurity begin
CheckSecurity(2)
' EditorialCatRecord PageSecurity end
'===============================

'===============================
' EditorialCatRecord Open Event begin
' EditorialCatRecord Open Event end
'===============================

'===============================
' EditorialCatRecord OpenAnyPage Event begin
' EditorialCatRecord OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' EditorialCatRecord Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
seditorial_categoriesErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "editorial_categories"
    editorial_categoriesAction(sAction)
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
editorial_categories_Show
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

' EditorialCatRecord Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' EditorialCatRecord Close Event begin
' EditorialCatRecord Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub editorial_categoriesAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKeditorial_cat_id : pPKeditorial_cat_id = ""
  Dim fldeditorial_cat_name : fldeditorial_cat_name = ""
'-------------------------------

'-------------------------------
' editorial_categories Action begin
'-------------------------------
  sActionFileName = "EditorialCatGrid.asp"

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' editorial_categories BeforeCancel Event begin
' editorial_categories BeforeCancel Event end
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
    pPKeditorial_cat_id = GetParam("PK_editorial_cat_id")
    if IsEmpty(pPKeditorial_cat_id) then exit sub
    sWhere = "editorial_cat_id=" & ToSQL(pPKeditorial_cat_id, "Number")
  end if
'-------------------------------


'-------------------------------
' Load all form fields into variables
'-------------------------------
  fldeditorial_cat_name = GetParam("editorial_cat_name")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
'-------------------------------
' editorial_categories Check Event begin
' editorial_categories Check Event end
'-------------------------------
    If len(seditorial_categoriesErr) > 0 then
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
' editorial_categories Insert Event begin
' editorial_categories Insert Event end
'-------------------------------
      sSQL = "insert into editorial_categories (" & _
          "[editorial_cat_name])" & _
          " values (" & _
          ToSQL(fldeditorial_cat_name, "Text") & _
          ")"
    case "update"
'-------------------------------
' editorial_categories Update Event begin
' editorial_categories Update Event end
'-------------------------------
      sSQL = "update editorial_categories set " & _
        "[editorial_cat_name]=" & ToSQL(fldeditorial_cat_name, "Text")
      sSQL = sSQL & " where " & sWhere
    case "delete"
'-------------------------------
' editorial_categories Delete Event begin
' editorial_categories Delete Event end
'-------------------------------
      sSQL = "delete from editorial_categories where " & sWhere
  end select
'-------------------------------
'-------------------------------
' editorial_categories BeforeExecute Event begin
' editorial_categories BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(seditorial_categoriesErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  seditorial_categoriesErr = ProcessError
  on error goto 0
  if len(seditorial_categoriesErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' editorial_categories Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub editorial_categories_Show()
'-------------------------------
' editorial_categories Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Editorial Categories"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if seditorial_categoriesErr = "" then
    fldeditorial_cat_id = GetParam("editorial_cat_id")
    peditorial_cat_id = GetParam("editorial_cat_id")
    SetVar "editorial_categoriesError", ""
  else
    fldeditorial_cat_id = GetParam("editorial_cat_id")
    fldeditorial_cat_name = GetParam("editorial_cat_name")
    peditorial_cat_id = GetParam("PK_editorial_cat_id")
    SetVar "seditorial_categoriesErr", seditorial_categoriesErr
    SetVar "FormTitle", sFormTitle
    Parse "editorial_categoriesError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

'-------------------------------

'-------------------------------
' Build WHERE statement

  if IsEmpty(peditorial_cat_id) then bPK = False
  
  sWhere = sWhere & "editorial_cat_id=" & ToSQL(peditorial_cat_id, "Number")
  SetVar "PK_editorial_cat_id", peditorial_cat_id
'-------------------------------
'-------------------------------
' editorial_categories Open Event begin
' editorial_categories Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from editorial_categories where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "editorial_categories") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldeditorial_cat_id = GetValue(rs, "editorial_cat_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if seditorial_categoriesErr = "" then
      fldeditorial_cat_name = GetValue(rs, "editorial_cat_name")
    end if
    SetVar "editorial_categoriesInsert", ""
    Parse "editorial_categoriesEdit", False
'-------------------------------
' editorial_categories ShowEdit Event begin
' editorial_categories ShowEdit Event end
'-------------------------------
  else
    if seditorial_categoriesErr = "" then
      fldeditorial_cat_id = ToHTML(GetParam("editorial_cat_id"))
    end if
    SetVar "editorial_categoriesEdit", ""
    Parse "editorial_categoriesInsert", False
'-------------------------------
' editorial_categories ShowInsert Event begin
' editorial_categories ShowInsert Event end
'-------------------------------
  end if
  Parse "editorial_categoriesCancel", false
'-------------------------------
' editorial_categories Show Event begin
' editorial_categories Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "editorial_cat_id", ToHTML(fldeditorial_cat_id)
      SetVar "editorial_cat_name", ToHTML(fldeditorial_cat_name)
  Parse "Formeditorial_categories", False

'-------------------------------
' editorial_categories Close Event begin
' editorial_categories Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' editorial_categories Show end
'-------------------------------
End Sub
'===============================
%>