<%
'
'    Filename: BookMaint.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' BookMaint CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' BookMaint CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "BookMaint.asp"
sTemplateFileName = "BookMaint.html"
'===============================


'===============================
' BookMaint PageSecurity begin
CheckSecurity(2)
' BookMaint PageSecurity end
'===============================

'===============================
' BookMaint Open Event begin
' BookMaint Open Event end
'===============================

'===============================
' BookMaint OpenAnyPage Event begin
' BookMaint OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' BookMaint Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sBookErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Book"
    BookAction(sAction)
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
Book_Show
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

' BookMaint Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' BookMaint Close Event begin
' BookMaint Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub BookAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sParams : sParams = "?"
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKitem_id : pPKitem_id = ""
  Dim fldname : fldname = ""
  Dim fldauthor : fldauthor = ""
  Dim fldcategory_id : fldcategory_id = ""
  Dim fldprice : fldprice = ""
  Dim fldproduct_url : fldproduct_url = ""
  Dim fldimage_url : fldimage_url = ""
  Dim fldnotes : fldnotes = ""
  Dim fldis_recommended : fldis_recommended = ""
'-------------------------------

'-------------------------------
' Book Action begin
'-------------------------------
  sActionFileName = "AdminBooks.asp"
  sParams = sParams & "category_id=" & ToURL(GetParam("Trn_category_id"))

'-------------------------------
' CANCEL action
'-------------------------------
  if sAction = "cancel" then

'-------------------------------
' Book BeforeCancel Event begin
' Book BeforeCancel Event end
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
    pPKitem_id = GetParam("PK_item_id")
    if IsEmpty(pPKitem_id) then exit sub
    sWhere = "item_id=" & ToSQL(pPKitem_id, "Number")
  end if
'-------------------------------


'-------------------------------
' Load all form fields into variables
'-------------------------------
  fldname = GetParam("name")
  fldauthor = GetParam("author")
  fldcategory_id = GetParam("category_id")
  fldprice = GetParam("price")
  fldproduct_url = GetParam("product_url")
  fldimage_url = GetParam("image_url")
  fldnotes = GetParam("notes")
  fldis_recommended = getCheckBoxValue(GetParam("is_recommended"), "1", "0", "Number")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldname) then
      sBookErr = sBookErr & "The value in field Title is required.<br>"
    end if
    if IsEmpty(fldcategory_id) then
      sBookErr = sBookErr & "The value in field Category is required.<br>"
    end if
    if IsEmpty(fldprice) then
      sBookErr = sBookErr & "The value in field Price is required.<br>"
    end if
    if not isNumeric(fldcategory_id) then
      sBookErr = sBookErr & "The value in field Category is incorrect.<br>"
    end if
    if not isNumeric(fldprice) then
      sBookErr = sBookErr & "The value in field Price is incorrect.<br>"
    end if
'-------------------------------
' Book Check Event begin
' Book Check Event end
'-------------------------------
    If len(sBookErr) > 0 then
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
' Book Insert Event begin
' Book Insert Event end
'-------------------------------
      sSQL = "insert into items (" & _
          "[name]," & _
          "[author]," & _
          "[category_id]," & _
          "[price]," & _
          "[product_url]," & _
          "[image_url]," & _
          "[notes]," & _
          "[is_recommended])" & _
          " values (" & _
          ToSQL(fldname, "Text") & "," & _
          ToSQL(fldauthor, "Text") & "," & _
          ToSQL(fldcategory_id, "Number") & "," & _
          ToSQL(fldprice, "Number") & "," & _
          ToSQL(fldproduct_url, "Text") & "," & _
          ToSQL(fldimage_url, "Text") & "," & _
          ToSQL(fldnotes, "Text") & "," & _
          fldis_recommended & _
          ")"
    case "update"
'-------------------------------
' Book Update Event begin
' Book Update Event end
'-------------------------------
      sSQL = "update items set " & _
        "[name]=" & ToSQL(fldname, "Text") & _
        ",[author]=" & ToSQL(fldauthor, "Text") & _
        ",[category_id]=" & ToSQL(fldcategory_id, "Number") & _
        ",[price]=" & ToSQL(fldprice, "Number") & _
        ",[product_url]=" & ToSQL(fldproduct_url, "Text") & _
        ",[image_url]=" & ToSQL(fldimage_url, "Text") & _
        ",[notes]=" & ToSQL(fldnotes, "Text") & _
        ",[is_recommended]=" & fldis_recommended
      sSQL = sSQL & " where " & sWhere
    case "delete"
'-------------------------------
' Book Delete Event begin
' Book Delete Event end
'-------------------------------
      sSQL = "delete from items where " & sWhere
  end select
'-------------------------------
'-------------------------------
' Book BeforeExecute Event begin
' Book BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sBookErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sBookErr = ProcessError
  on error goto 0
  if len(sBookErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName & sParams
'-------------------------------
' Book Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Book_Show()
'-------------------------------
' Book Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Book"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sBookErr = "" then
    fldcategory_id = GetParam("category_id")
    flditem_id = GetParam("item_id")
    SetVar "Trn_category_id", GetParam("category_id")
    pitem_id = GetParam("item_id")
    SetVar "BookError", ""
  else
    flditem_id = GetParam("item_id")
    fldname = GetParam("name")
    fldauthor = GetParam("author")
    fldcategory_id = GetParam("category_id")
    fldprice = GetParam("price")
    fldproduct_url = GetParam("product_url")
    fldimage_url = GetParam("image_url")
    fldnotes = GetParam("notes")
    fldis_recommended = getCheckBoxValue(GetParam("is_recommended"), "1", "0", "Number")
    SetVar "Trn_category_id", GetParam("Trn_category_id")
    pitem_id = GetParam("PK_item_id")
    SetVar "sBookErr", sBookErr
    SetVar "FormTitle", sFormTitle
    Parse "BookError", False
  end if
'-------------------------------

'-------------------------------
' Load all form fields

'-------------------------------

'-------------------------------
' Build WHERE statement

  if IsEmpty(pitem_id) then bPK = False
  
  sWhere = sWhere & "item_id=" & ToSQL(pitem_id, "Number")
  SetVar "PK_item_id", pitem_id
'-------------------------------
'-------------------------------
' Book Open Event begin
' Book Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from items where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Book") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    flditem_id = GetValue(rs, "item_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sBookErr = "" then
      fldname = GetValue(rs, "name")
      fldauthor = GetValue(rs, "author")
      fldcategory_id = GetValue(rs, "category_id")
      fldprice = GetValue(rs, "price")
      fldproduct_url = GetValue(rs, "product_url")
      fldimage_url = GetValue(rs, "image_url")
      fldnotes = GetValue(rs, "notes")
      fldis_recommended = GetValue(rs, "is_recommended")
    end if
    SetVar "BookInsert", ""
    Parse "BookEdit", False
'-------------------------------
' Book ShowEdit Event begin
' Book ShowEdit Event end
'-------------------------------
  else
    if sBookErr = "" then
      flditem_id = ToHTML(GetParam("item_id"))
      fldcategory_id = ToHTML(GetParam("category_id"))
      fldis_recommended= "0"
    end if
    SetVar "BookEdit", ""
    Parse "BookInsert", False
'-------------------------------
' Book ShowInsert Event begin
' Book ShowInsert Event end
'-------------------------------
  end if
  Parse "BookCancel", false
'-------------------------------
' Book Show Event begin
' Book Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "item_id", ToHTML(flditem_id)
      SetVar "name", ToHTML(fldname)
      SetVar "author", ToHTML(fldauthor)
      SetVar "BookLBcategory_id", ""
      openrs rscategory_id, "select category_id, name from categories order by 2"
      while not rscategory_id.EOF
        SetVar "ID", GetValue(rscategory_id, 0) : SetVar "Value", GetValue(rscategory_id, 1)
        if cstr(GetValue(rscategory_id, 0)) = cstr(fldcategory_id) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
        Parse "BookLBcategory_id", True
        rscategory_id.MoveNext
      wend
      set rscategory_id = nothing
    
      SetVar "price", ToHTML(fldprice)
      SetVar "product_url", ToHTML(fldproduct_url)
      SetVar "image_url", ToHTML(fldimage_url)
      SetVar "notes", ToHTML(fldnotes)
  if (LCase(fldis_recommended) = LCase("1")) then
    SetVar "is_recommended_CHECKED", "CHECKED"
  else
    SetVar "is_recommended_CHECKED", ""
  end if

  Parse "FormBook", False

'-------------------------------
' Book Close Event begin
' Book Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Book Show end
'-------------------------------
End Sub
'===============================
%>