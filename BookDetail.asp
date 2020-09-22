<%
'
'    Filename: BookDetail.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' BookDetail CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' BookDetail CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "BookDetail.asp"
sTemplateFileName = "BookDetail.html"
'===============================


'===============================
' BookDetail PageSecurity begin
CheckSecurity(1)
' BookDetail PageSecurity end
'===============================

'===============================
' BookDetail Open Event begin
' BookDetail Open Event end
'===============================

'===============================
' BookDetail OpenAnyPage Event begin
' BookDetail OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' BookDetail Show begin
'===============================
' Perform the form's action
'-------------------------------
' Initialize error variables
'-------------------------------
sDetailErr = ""
sOrderErr = ""
sRatingErr = ""

'-------------------------------
' Select the FormAction
'-------------------------------
Select Case sForm
  Case "Detail"
    DetailAction(sAction)
  Case "Order"
    OrderAction(sAction)
  Case "Rating"
    RatingAction(sAction)
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
Detail_Show
Order_Show
Rating_Show
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

' BookDetail Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' BookDetail Close Event begin
' BookDetail Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub DetailAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sParams : sParams = "?"
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKitem_id : pPKitem_id = ""
'-------------------------------

'-------------------------------
' Detail Action begin
'-------------------------------
  sActionFileName = "ShoppingCart.asp"
  sParams = sParams & "item_id=" & ToURL(GetParam("Trn_item_id"))

'-------------------------------
' Load all form fields into variables
'-------------------------------
'-------------------------------
' Detail BeforeExecute Event begin
' Detail BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sDetailErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sDetailErr = ProcessError
  on error goto 0
  if len(sDetailErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName & sParams
'-------------------------------
' Detail Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Detail_Show()
'-------------------------------
' Detail Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Book Detail"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sDetailErr = "" then
    flditem_id = GetParam("item_id")
    SetVar "Trn_item_id", GetParam("item_id")
    pitem_id = GetParam("item_id")
    SetVar "DetailError", ""
  else
    flditem_id = GetParam("item_id")
    SetVar "Trn_item_id", GetParam("Trn_item_id")
    pitem_id = GetParam("PK_item_id")
    SetVar "sDetailErr", sDetailErr
    SetVar "FormTitle", sFormTitle
    Parse "DetailError", False
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
' Detail Open Event begin
' Detail Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from items where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Detail") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    fldauthor = GetValue(rs, "author")
    fldcategory_id = GetValue(rs, "category_id")
    fldimage_url_URLLink = GetValue(rs, "product_url")
    fldimage_url = GetValue(rs, "image_url")
    flditem_id = GetValue(rs, "item_id")
    fldname = GetValue(rs, "name")
    fldnotes = GetValue(rs, "notes")
    fldprice = GetValue(rs, "price")
    fldproduct_url_URLLink = GetValue(rs, "product_url")
    fldproduct_url = GetValue(rs, "product_url")
    SetVar "DetailDelete", ""
    SetVar "DetailUpdate", ""
    SetVar "DetailInsert", ""
'-------------------------------
' Detail ShowEdit Event begin
' Detail ShowEdit Event end
'-------------------------------
  else
    if sDetailErr = "" then
      flditem_id = ToHTML(GetParam("item_id"))
    end if
    SetVar "DetailEdit", ""
    SetVar "DetailInsert", ""
'-------------------------------
' Detail ShowInsert Event begin
' Detail ShowInsert Event end
'-------------------------------
  end if
  SetVar "DetailCancel", ""

'-------------------------------
' Set lookup fields
'-------------------------------
  fldcategory_id = DLookUp("categories", "name", "category_id=" & ToSQL(fldcategory_id, "Number"))
  if sDetailErr = "" then
'-------------------------------
' Detail Show Event begin
fldimage_url="<img border=""0"" src=""" & fldimage_url & """/>"
fldproduct_url="Review this book on Amazon.com"
' Detail Show Event end
'-------------------------------
  end if

'-------------------------------
' Show form field
'-------------------------------
      SetVar "item_id", ToHTML(flditem_id)
      SetVar "name", ToHTML(fldname)
      SetVar "author", ToHTML(fldauthor)
      SetVar "category_id", ToHTML(fldcategory_id)
      SetVar "price", ToHTML(fldprice)
      SetVar "image_url", fldimage_url
      SetVar "image_url_URLLink", fldimage_url_URLLink
      SetVar "notes", fldnotes
      SetVar "product_url", ToHTML(fldproduct_url)
      SetVar "product_url_URLLink", fldproduct_url_URLLink
  Parse "FormDetail", False

'-------------------------------
' Detail Close Event begin
' Detail Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Detail Show end
'-------------------------------
End Sub
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub OrderAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKorder_id : pPKorder_id = ""
  Dim fldquantity : fldquantity = ""
  Dim flditem_id : flditem_id = ""
'-------------------------------

'-------------------------------
' Order Action begin
'-------------------------------
  sActionFileName = "ShoppingCart.asp"

'-------------------------------
' Load all form fields into variables
'-------------------------------
  fldUserID = Session("UserID")
  fldquantity = GetParam("quantity")
  flditem_id = GetParam("item_id")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldquantity) then
      sOrderErr = sOrderErr & "The value in field Quantity is required.<br>"
    end if
    if not isNumeric(fldquantity) then
      sOrderErr = sOrderErr & "The value in field Quantity is incorrect.<br>"
    end if
    if not isNumeric(flditem_id) then
      sOrderErr = sOrderErr & "The value in field item_id is incorrect.<br>"
    end if
'-------------------------------
' Order Check Event begin
' Order Check Event end
'-------------------------------
    If len(sOrderErr) > 0 then
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
' Order Insert Event begin
' Order Insert Event end
'-------------------------------
      sSQL = "insert into orders (" & _
          "[member_id]," & _
          "[quantity]," & _
          "[item_id])" & _
          " values (" & _
          ToSQL(fldUserID, "Number") & "," & _
          ToSQL(fldquantity, "Number") & "," & _
          ToSQL(flditem_id, "Number") & _
          ")"
  end select
'-------------------------------
'-------------------------------
' Order BeforeExecute Event begin
' Order BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sOrderErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sOrderErr = ProcessError
  on error goto 0
  if len(sOrderErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName
'-------------------------------
' Order Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Order_Show()
'-------------------------------
' Order Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = ""
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sOrderErr = "" then
    flditem_id = GetParam("item_id")
    porder_id = GetParam("order_id")
    SetVar "OrderError", ""
  else
    fldorder_id = GetParam("order_id")
    fldquantity = GetParam("quantity")
    flditem_id = GetParam("item_id")
    porder_id = GetParam("PK_order_id")
    SetVar "sOrderErr", sOrderErr
    SetVar "FormTitle", sFormTitle
    Parse "OrderError", False
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
' Order Open Event begin
' Order Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from orders where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Order") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    flditem_id = GetValue(rs, "item_id")
    fldorder_id = GetValue(rs, "order_id")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sOrderErr = "" then
      fldquantity = GetValue(rs, "quantity")
    end if
    SetVar "OrderDelete", ""
    SetVar "OrderUpdate", ""
    SetVar "OrderInsert", ""
'-------------------------------
' Order ShowEdit Event begin
' Order ShowEdit Event end
'-------------------------------
  else
    if sOrderErr = "" then
      flditem_id = ToHTML(GetParam("item_id"))
      fldquantity= "1"
    end if
    SetVar "OrderEdit", ""
    Parse "OrderInsert", False
'-------------------------------
' Order ShowInsert Event begin
' Order ShowInsert Event end
'-------------------------------
  end if
  SetVar "OrderCancel", ""
'-------------------------------
' Order Show Event begin
' Order Show Event end
'-------------------------------

'-------------------------------
' Show form field
'-------------------------------
      SetVar "order_id", ToHTML(fldorder_id)
      SetVar "quantity", ToHTML(fldquantity)
      SetVar "item_id", ToHTML(flditem_id)
  Parse "FormOrder", False

'-------------------------------
' Order Close Event begin
' Order Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Order Show end
'-------------------------------
End Sub
'===============================
'===============================
' Action of the Record Form
'-------------------------------
Sub RatingAction(sAction)
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim bExecSQL: bExecSQL = true
  Dim sActionFileName : sActionFileName = ""
  Dim sParams : sParams = "?"
  Dim sWhere : sWhere = "" 
  Dim bErr : bErr = False
  Dim pPKitem_id : pPKitem_id = ""
  Dim fldrating : fldrating = ""
  Dim fldrating_count : fldrating_count = ""
'-------------------------------

'-------------------------------
' Rating Action begin
'-------------------------------
  sActionFileName = "BookDetail.asp"
  sParams = sParams & "item_id=" & ToURL(GetParam("Trn_item_id"))

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
  fldrating = GetParam("rating")
  fldrating_count = GetParam("rating_count")

'-------------------------------
' Validate fields
'-------------------------------
  if sAction = "insert" or sAction = "update" then
    if IsEmpty(fldrating) then
      sRatingErr = sRatingErr & "The value in field Your Rating is required.<br>"
    end if
    if not isNumeric(fldrating) then
      sRatingErr = sRatingErr & "The value in field Your Rating is incorrect.<br>"
    end if
    if not isNumeric(fldrating_count) then
      sRatingErr = sRatingErr & "The value in field rating_count is incorrect.<br>"
    end if
'-------------------------------
' Rating Check Event begin
' Rating Check Event end
'-------------------------------
    If len(sRatingErr) > 0 then
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
' Rating Update Event begin
sSQL="update items set rating=rating+" & getparam("rating") & ", rating_count=rating_count+1 where item_id=" & getparam("item_id")
' Rating Update Event end
'-------------------------------
      if sSQL = "" then
      sSQL = "update items set " & _
        "[rating]=" & ToSQL(fldrating, "Number") & _
        ",[rating_count]=" & ToSQL(fldrating_count, "Number")
      sSQL = sSQL & " where " & sWhere
      end if
  end select
'-------------------------------
'-------------------------------
' Rating BeforeExecute Event begin
' Rating BeforeExecute Event end
'-------------------------------

'-------------------------------
' Execute SQL statement
'-------------------------------
if len(sRatingErr) > 0 then Exit Sub
  on error resume next
  if bExecSQL then 
    cn.execute sSQL
  end if
  sRatingErr = ProcessError
  on error goto 0
  if len(sRatingErr) > 0 then Exit Sub
  cn.Close
  Set cn = Nothing
  response.redirect sActionFileName & sParams
'-------------------------------
' Rating Action end
'-------------------------------
end sub
'===============================

'===============================
' Display Record Form
'-------------------------------
Sub Rating_Show()
'-------------------------------
' Rating Show begin
'-------------------------------
  Dim sWhere : sWhere = ""
  Dim sFormTitle: sFormTitle = "Rating"
  Dim bPK : bPK = True

'-------------------------------
' Load primary key and form parameters
'-------------------------------
  if sRatingErr = "" then
    flditem_id = GetParam("item_id")
    SetVar "Trn_item_id", GetParam("item_id")
    pitem_id = GetParam("item_id")
    SetVar "RatingError", ""
  else
    flditem_id = GetParam("item_id")
    fldrating = GetParam("rating")
    fldrating_count = GetParam("rating_count")
    SetVar "Trn_item_id", GetParam("Trn_item_id")
    pitem_id = GetParam("PK_item_id")
    SetVar "sRatingErr", sRatingErr
    SetVar "FormTitle", sFormTitle
    Parse "RatingError", False
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
' Rating Open Event begin
' Rating Open Event end
'-------------------------------

  SetVar "FormTitle", sFormTitle

'-------------------------------
' Build SQL statement and open recordset
'-------------------------------
  sSQL = "select * from items where " & sWhere
  openrs rs, sSQL
  bIsUpdateMode = (bPK and not(sAction = "insert" and sForm = "Rating") and not rs.eof)
'-------------------------------

'-------------------------------
' Load all fields into variables from recordset or input parameters
'-------------------------------
  if bIsUpdateMode then
    flditem_id = GetValue(rs, "item_id")
    fldrating_view = GetValue(rs, "rating")
    fldrating_count_view = GetValue(rs, "rating_count")
    fldrating_count = GetValue(rs, "rating_count")

'-------------------------------
' Load data from recordset when form displayed first time
'-------------------------------
    if sRatingErr = "" then
      fldrating = GetValue(rs, "rating")
    end if
    SetVar "RatingDelete", ""
    SetVar "RatingInsert", ""
    Parse "RatingEdit", False
'-------------------------------
' Rating ShowEdit Event begin
' Rating ShowEdit Event end
'-------------------------------
  else
    if sRatingErr = "" then
      flditem_id = ToHTML(GetParam("item_id"))
    end if
    SetVar "RatingEdit", ""
    SetVar "RatingInsert", ""
'-------------------------------
' Rating ShowInsert Event begin
' Rating ShowInsert Event end
'-------------------------------
  end if
  SetVar "RatingCancel", ""
  if sRatingErr = "" then
'-------------------------------
' Rating Show Event begin
if fldrating_view=0 then
  fldrating_view="Not rated yet"
  fldrating_count_view=""
else
  fldrating_view="<img src=images/" & round(fldrating/fldrating_count) & "stars.gif>"
end if
' Rating Show Event end
'-------------------------------
  end if

'-------------------------------
' Show form field
'-------------------------------
      SetVar "item_id", ToHTML(flditem_id)
      SetVar "rating_view", fldrating_view
      SetVar "rating_count_view", ToHTML(fldrating_count_view)
      SetVar "RatingLBrating", ""
      LOV = Split("1;Deficient;2;Regular;3;Good;4;Very Good;5;Excellent", ";")
      if (ubound(LOV) mod 2) = 1 then
        for i = 0 to ubound(LOV) step 2
          SetVar "ID", LOV(i) : SetVar "Value", LOV(i+1)
          if cstr(LOV(i)) = cstr(fldrating) then SetVar "Selected", "SELECTED" else SetVar "Selected", ""
          Parse "RatingLBrating", True
        next
      end if
    
      SetVar "rating_count", ToHTML(fldrating_count)
  Parse "FormRating", False

'-------------------------------
' Rating Close Event begin
' Rating Close Event end
'-------------------------------

Set rs = Nothing

'-------------------------------
' Rating Show end
'-------------------------------
End Sub
'===============================
%>