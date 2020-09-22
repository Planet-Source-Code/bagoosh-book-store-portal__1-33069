<%
'
'    Filename: EditorialsRecord.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' EditorialsRecord CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' EditorialsRecord CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "EditorialsRecord.asp"
sTemplateFileName = "EditorialsRecord.html"
'===============================


'===============================
' EditorialsRecord PageSecurity begin
CheckSecurity(2)
' EditorialsRecord PageSecurity end
'===============================

'===============================
' EditorialsRecord Open Event begin
' EditorialsRecord Open Event end
'===============================

'===============================
' EditorialsRecord OpenAnyPage Event begin
' EditorialsRecord OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' EditorialsRecord Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
seditorialsErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "editorials"
    editorialsAction(sAction)
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
editorials_Show
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

' EditorialsRecord Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' EditorialsRecord Close Event begin
' EditorialsRecord Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub editorialsAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKarticle_id : pPKarticle_id = ""
  Dim fldarticle_desc : fldarticle_desc = ""
  Dim fldarticle_title : fldarticle_title = ""
  Dim fldeditorial_cat_id : fldeditorial_cat_id = ""
  Dim flditem_id : flditem_id = ""
'-------------------------------

'-------------------------------
' editorials Action begin
'-------------------------------
  sActionFileName = "EditorialsGrid.asp"

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' editorials BeforeCancel Event begin
' editorials BeforeCancel Event end
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
    pPKarticle_id = GetParam("PK_article_id")
    if IsEmpty(pPKarticle_id) then exit sub
    sWhere = "article_id=" & ToSQL(pPKarticle_id, "Number")
  end if
'-------------------------------


'-------------------------------
' Load all form fields into variables
'-------------------------------
  fldarticle_desc = GetParam("article_desc")
  fldarticle_title = GetParam("article_title")
  fldeditorial_cat_id = GetParam("editorial_cat_id")
  flditem_id = GetParam("item_id")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldeditorial_cat_id) then
      seditorialsErr = seditorialsErr & "The value in field Editorial Category is required.<br>"
    end if
    if not isNumeric(fldeditorial_cat_id) then
      seditorialsErr = seditorialsErr & "The value in field Editorial Category is incorrect.<br>"
    end if
    if not isNumeric(flditem_id) then
      seditorialsErr = seditorialsErr & "The value in field Item is incorrect.<br>"
    end if
'-------------------------------
' editorials Check Event begin
' editorials Check Event end
'-------------------------------
    If len(seditorialsErr) > 0 then
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
' editorials Insert Event begin
' editorials Insert Event end
'-------------------------------
      sSQL = "insert into editorials (" & _
          "[article_desc]," & _
          "[article_title]," & _
          "[editorial_cat_id]," & _
          "[item_id])" & _
          " values (" & _
          ToSQL(fldarticle_desc, "Text") & "," & _
          ToSQL(fldarticle_title, "Text") & "," & _
          ToSQL(fldeditorial_cat_id, "Number") & "," & _
          ToSQL(flditem_id, "Number") & _
          ")"
    case "update"
'-------------------------------
' editorials Update Event begin
' editorials Update Event end
'-------------------------------
      sSQL = "update editorials set " & _
        "[article_desc]=" & ToSQL(fldarticle_desc, "Text") & _
        ",[article_title]=" & ToSQL(fldarticle_title, "Text") & _
        ",[editorial_cat_id]=" & ToSQL(fldeditorial_cat_id, "Number") & _
        ",[item_id]=" & ToSQL(flditem_id, "Number")
      sSQL = sSQL & " where " & sWhere
    case "delete"
'-------------------------------
' editorials Delete Event begin
' editorials Delete Event end
'-------------------------------
      sSQL = "delete from editorials where " & sWhere
  end select
'-------------------------------
'-------------------------------
' editorials BeforeExecute Event begin
' editorials BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(seditorialsErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  seditorialsErr = ProcessError
  on error goto 0
  if len(seditorialsErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' editorials Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub editorials_Show()
'-------------------------------
' editorials Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Editorial"
  Dim bPK : bPK = True
  Dim sitem_idDisplayValue: sitem_idDisplayValue = ""

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if seditorialsErr = "" then
    fldarticle_id = GetParam("article_id")
    particle_id = GetParam("article_id")
    SetVar "editorialsError", ""
  else
    fldarticle_id = GetParam("article_id")
    fldarticle_desc = GetParam("article_desc")
    fldarticle_title = GetParam("article_title")
    fldeditorial_cat_id = GetParam("editorial_cat_id")
    flditem_id = GetParam("item_id")
    particle_id = GetParam("PK_article_id")
    SetVar "seditorialsErr", seditorialsErr
    SetVar "FormTitle", sFormTitle
    Parse "editorialsError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

'-------------------------------

'-------------------------------
' Build WHERE statement

  if IsEmpty(particle_id) then bPK = False
  
  sWhere = sWhere & "article_id=" & ToSQL(particle_id, "Number")
  SetVar "PK_article_id", particle_id
'-------------------------------
'-------------------------------
' editorials Open Event begin
' editorials Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from editorials where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "editorials") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldarticle_id = GetValue(rs, "article_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if seditorialsErr = "" then
      fldarticle_desc = GetValue(rs, "article_desc")
      fldarticle_title = GetValue(rs, "article_title")
      fldeditorial_cat_id = GetValue(rs, "editorial_cat_id")
      flditem_id = GetValue(rs, "item_id")
    end if
    SetVar "editorialsInsert", ""
    Parse "editorialsEdit", False
'-------------------------------
' editorials ShowEdit Event begin
' editorials ShowEdit Event end
'-------------------------------
  else
    if seditorialsErr = "" then
      fldarticle_id = ToHTML(GetParam("article_id"))
    end if
    SetVar "editorialsEdit", ""
    Parse "editorialsInsert", False
'-------------------------------
' editorials ShowInsert Event begin
' editorials ShowInsert Event end
'-------------------------------
  end if
  Parse "editorialsCancel", false
'-------------------------------
' editorials Show Event begin
' editorials Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "article_id", ToHTML(fldarticle_id)
      SetVar "article_desc", ToHTML(fldarticle_desc)
      SetVar "article_title", ToHTML(fldarticle_title)
      SetVar "editorialsLBeditorial_cat_id", ""
      openrs rseditorial_cat_id, "select editorial_cat_id, editorial_cat_name from editorial_categories order by 2"
      while not rseditorial_cat_id.EOF
        SetVar "ID", GetValue(rseditorial_cat_id, 0) : SetVar "Value", GetValue(rseditorial_cat_id, 1)
        if cstr(GetValue(rseditorial_cat_id, 0)) = cstr(fldeditorial_cat_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "editorialsLBeditorial_cat_id", True
        rseditorial_cat_id.MoveNext
      wend
      set rseditorial_cat_id = nothing
    
      SetVar "editorialsLBitem_id", ""
      SetVar "Selected", ""
      SetVar "ID", ""
      SetVar "Value", sitem_idDisplayValue
      Parse "editorialsLBitem_id", True
      openrs rsitem_id, "select item_id, name from items order by 2"
      while not rsitem_id.EOF
        SetVar "ID", GetValue(rsitem_id, 0) : SetVar "Value", GetValue(rsitem_id, 1)
        if cstr(GetValue(rsitem_id, 0)) = cstr(flditem_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "editorialsLBitem_id", True
        rsitem_id.MoveNext
      wend
      set rsitem_id = nothing
    
  Parse "Formeditorials", False

'-------------------------------
' editorials Close Event begin
' editorials Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' editorials Show end
'-------------------------------
End Sub
'===============================
%>