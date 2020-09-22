<%
'
'    Filename: CategoriesRecord.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' CategoriesRecord CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' CategoriesRecord CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "CategoriesRecord.asp"
sTemplateFileName = "CategoriesRecord.html"
'===============================


'===============================
' CategoriesRecord PageSecurity begin
CheckSecurity(2)
' CategoriesRecord PageSecurity end
'===============================

'===============================
' CategoriesRecord Open Event begin
' CategoriesRecord Open Event end
'===============================

'===============================
' CategoriesRecord OpenAnyPage Event begin
' CategoriesRecord OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' CategoriesRecord Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sCategoriesErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Categories"
    CategoriesAction(sAction)
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
Categories_Show
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

' CategoriesRecord Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' CategoriesRecord Close Event begin
' CategoriesRecord Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub CategoriesAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKcategory_id : pPKcategory_id = ""
  Dim fldname : fldname = ""
'-------------------------------

'-------------------------------
' Categories Action begin
'-------------------------------
  sActionFileName = "CategoriesGrid.asp"

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' Categories BeforeCancel Event begin
' Categories BeforeCancel Event end
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
    pPKcategory_id = GetParam("PK_category_id")
    if IsEmpty(pPKcategory_id) then exit sub
    sWhere = "category_id=" & ToSQL(pPKcategory_id, "Number")
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
      sCategoriesErr = sCategoriesErr & "The value in field Name is required.<br>"
    end if
'-------------------------------
' Categories Check Event begin
' Categories Check Event end
'-------------------------------
    If len(sCategoriesErr) > 0 then
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
' Categories Insert Event begin
' Categories Insert Event end
'-------------------------------
      sSQL = "insert into categories (" & _
          "[name])" & _
          " values (" & _
          ToSQL(fldname, "Text") & _
          ")"
    case "update"
'-------------------------------
' Categories Update Event begin
' Categories Update Event end
'-------------------------------
      sSQL = "update categories set " & _
        "[name]=" & ToSQL(fldname, "Text")
      sSQL = sSQL & " where " & sWhere
    case "delete"
'-------------------------------
' Categories Delete Event begin
' Categories Delete Event end
'-------------------------------
      sSQL = "delete from categories where " & sWhere
  end select
'-------------------------------
'-------------------------------
' Categories BeforeExecute Event begin
' Categories BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sCategoriesErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sCategoriesErr = ProcessError
  on error goto 0
  if len(sCategoriesErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' Categories Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Categories_Show()
'-------------------------------
' Categories Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Categories"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sCategoriesErr = "" then
    fldcategory_id = GetParam("category_id")
    pcategory_id = GetParam("category_id")
    SetVar "CategoriesError", ""
  else
    fldcategory_id = GetParam("category_id")
    fldname = GetParam("name")
    pcategory_id = GetParam("PK_category_id")
    SetVar "sCategoriesErr", sCategoriesErr
    SetVar "FormTitle", sFormTitle
    Parse "CategoriesError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

'-------------------------------

'-------------------------------
' Build WHERE statement

  if IsEmpty(pcategory_id) then bPK = False
  
  sWhere = sWhere & "category_id=" & ToSQL(pcategory_id, "Number")
  SetVar "PK_category_id", pcategory_id
'-------------------------------
'-------------------------------
' Categories Open Event begin
' Categories Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from categories where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Categories") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldcategory_id = GetValue(rs, "category_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sCategoriesErr = "" then
      fldname = GetValue(rs, "name")
    end if
    SetVar "CategoriesInsert", ""
    Parse "CategoriesEdit", False
'-------------------------------
' Categories ShowEdit Event begin
' Categories ShowEdit Event end
'-------------------------------
  else
    if sCategoriesErr = "" then
      fldcategory_id = ToHTML(GetParam("category_id"))
    end if
    SetVar "CategoriesEdit", ""
    Parse "CategoriesInsert", False
'-------------------------------
' Categories ShowInsert Event begin
' Categories ShowInsert Event end
'-------------------------------
  end if
  Parse "CategoriesCancel", false
'-------------------------------
' Categories Show Event begin
' Categories Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "category_id", ToHTML(fldcategory_id)
      SetVar "name", ToHTML(fldname)
  Parse "FormCategories", False

'-------------------------------
' Categories Close Event begin
' Categories Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Categories Show end
'-------------------------------
End Sub
'===============================
%>