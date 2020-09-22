<%
'
'    Filename: CardTypesGrid.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' CardTypesGrid CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' CardTypesGrid CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "CardTypesGrid.asp"
sTemplateFileName = "CardTypesGrid.html"
'===============================


'===============================
' CardTypesGrid PageSecurity begin
CheckSecurity(2)
' CardTypesGrid PageSecurity end
'===============================

'===============================
' CardTypesGrid Open Event begin
' CardTypesGrid Open Event end
'===============================

'===============================
' CardTypesGrid OpenAnyPage Event begin
' CardTypesGrid OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' CardTypesGrid Show begin

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

' CardTypesGrid Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' CardTypesGrid Close Event begin
' CardTypesGrid Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Grid Form
'-------------------------------
Sub CardTypes_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Card Types"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0
  Dim sActionFileName : sActionFileName = "CardTypesRecord.asp"

  SetVar "TransitParams", ""
  SetVar "FormParams", ""


  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  sOrder = " order by c.name Asc"
  iSort = GetParam("FormCardTypes_Sorting")
  iSorted = GetParam("FormCardTypes_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "FormCardTypes_Sorting=" & iSort & "&FormCardTypes_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "FormCardTypes_Sorting=" & iSort & "&FormCardTypes_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by c.[name]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [c].[card_type_id] as c_card_type_id, " & _
    "[c].[name] as c_name " & _
    " from [card_types] c "
'-------------------------------

'-------------------------------
' CardTypes Open Event begin
' CardTypes Open Event end
'-------------------------------

'-------------------------------
' Assemble full SQL statement
'-------------------------------
  sSQL = sSQL & sWhere & sOrder
'-------------------------------

SetVar "FormTitle", sFormTitle

'-------------------------------
' Process the link to the record page
'-------------------------------
  SetVar "FormAction", sActionFileName
'-------------------------------

'-------------------------------
' Process the parameters for sorting
'-------------------------------
  SetVar "SortParams", sSortParams
'-------------------------------

'-------------------------------
' Open the recordset
'-------------------------------
  openrs rs, sSQL
'-------------------------------

'-------------------------------
' Process empty recordset
'-------------------------------
  if rs.eof then
    set rs = nothing
    SetVar "DListCardTypes", ""
    Parse "CardTypesNoRecords", False
    Parse "FormCardTypes", False
    exit sub
  end if
'-------------------------------

'-------------------------------
' Initialize page counter and records per page
'-------------------------------
  iRecordsPerPage = 20
  iCounter = 0
'-------------------------------

'-------------------------------
' Display grid based on recordset
'-------------------------------
  while not rs.EOF  and iCounter < iRecordsPerPage
'-------------------------------
' Create field variables based on database fields
'-------------------------------
    fldname_URLLink = "CardTypesRecord.asp"
    fldname_card_type_id = GetValue(rs, "c_card_type_id")
    fldname = GetValue(rs, "c_name")
'-------------------------------
' CardTypes Show begin
'-------------------------------

'-------------------------------
' CardTypes Show Event begin
' CardTypes Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "name", ToHTML(fldname)
      SetVar "name_URLLink", fldname_URLLink
      SetVar "Prmname_card_type_id", ToURL(fldname_card_type_id)
    Parse "DListCardTypes", True

'-------------------------------
' CardTypes Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "CardTypesNoRecords", ""
  Parse "FormCardTypes", False

'-------------------------------
' CardTypes Close Event begin
' CardTypes Close Event end
'-------------------------------
End Sub
'===============================

%>