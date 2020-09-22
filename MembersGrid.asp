<%
'
'    Filename: MembersGrid.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'

'-------------------------------
' MembersGrid CustomIncludes begin
%>

<!-- #INCLUDE FILE="Common.asp" -->
<!-- #INCLUDE FILE="Header.asp" -->
<!-- #INCLUDE FILE="Footer.asp" -->

<%
' MembersGrid CustomIncludes end
'-------------------------------

'===============================
' Save Page and File Name available into variables
'-------------------------------
sFileName = "MembersGrid.asp"
sTemplateFileName = "MembersGrid.html"
'===============================


'===============================
' MembersGrid PageSecurity begin
CheckSecurity(2)
' MembersGrid PageSecurity end
'===============================

'===============================
' MembersGrid Open Event begin
' MembersGrid Open Event end
'===============================

'===============================
' MembersGrid OpenAnyPage Event begin
' MembersGrid OpenAnyPage Event end
'===============================

'===============================
'Save the name of the form and type of action into the variables
'-------------------------------
sAction = GetParam("FormAction")
sForm = GetParam("FormName")
'===============================

' MembersGrid Show begin

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
Search_Show
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

' MembersGrid Show end

'-------------------------------
' Destroy all object variables
'-------------------------------

' MembersGrid Close Event begin
' MembersGrid Close Event end

cn.Close
Set cn = Nothing
UnloadTemplate
'===============================

'===============================
' Display Search Form
'-------------------------------
Sub Search_Show()
  Dim sFormTitle: sFormTitle = ""
  Dim sActionFileName: sActionFileName = "MembersGrid.asp"

'-------------------------------
' Search Open Event begin
' Search Open Event end
'-------------------------------
      SetVar "FormTitle", sFormTitle
      SetVar "ActionPage", sActionFileName

'-------------------------------
' Set variables with search parameters
'-------------------------------
      fldname = GetParam("name")

'-------------------------------
' Search Show begin
'-------------------------------


'-------------------------------
' Search Show Event begin
' Search Show Event end
'-------------------------------
      SetVar "name", ToHTML(fldname)

'-------------------------------
' Search Show end
'-------------------------------

'-------------------------------
' Search Close Event begin
' Search Close Event end
'-------------------------------
      Parse "FormSearch", False
End Sub
'===============================


'===============================
' Display Grid Form
'-------------------------------
Sub Members_Show()
'-------------------------------
' Initialize variables  
'-------------------------------
  Dim rs
  Dim sWhere : sWhere = ""
  Dim sOrder : sOrder = ""
  Dim sSQL : sSQL = ""
  Dim sFormTitle: sFormTitle = "Members"
  Dim HasParam : HasParam = false
  Dim iSort : iSort = ""
  Dim iSorted : iSorted = ""
  Dim sDirection : sDirection = ""
  Dim sSortParams : sSortParams = ""
  Dim iRecordsPerPage : iRecordsPerPage = 20
  Dim iCounter : iCounter = 0
  Dim iPage : iPage = 0
  Dim bEof : bEof = False
  Dim sActionFileName : sActionFileName = "MembersRecord.asp"

  SetVar "TransitParams", "name=" & ToURL(GetParam("name")) & "&"
  SetVar "FormParams", "name=" & ToURL(GetParam("name")) & "&"

'-------------------------------
' Build WHERE statement
'-------------------------------
  pname = GetParam("name")
  if not isEmpty(pname) then
    HasParam = true
    sWhere = "m.[member_login] like '%" & replace(pname, "'", "''") & "%'" & " or " & "m.[first_name] like '%" & replace(pname, "'", "''") & "%'" & " or " & "m.[last_name] like '%" & replace(pname, "'", "''") & "%'"
  end if


  if HasParam then
    sWhere = " WHERE (" & sWhere & ")"
  end if
  
'-------------------------------
' Build ORDER BY statement
'-------------------------------
  sOrder = " order by m.member_login Asc"
  iSort = GetParam("FormMembers_Sorting")
  iSorted = GetParam("FormMembers_Sorted")
  sDirection = ""
  if IsEmpty(iSort) then
    SetVar "Form_Sorting", ""
  else
    if iSort = iSorted then 
      SetVar "Form_Sorting", ""
      sDirection = " DESC"
      sSortParams = "FormMembers_Sorting=" & iSort & "&FormMembers_Sorted=" & iSort & "&"
    else
      SetVar "Form_Sorting", iSort
      sDirection = " ASC"
      sSortParams = "FormMembers_Sorting=" & iSort & "&FormMembers_Sorted=" & "&"
    end if
    if iSort = 1 then sOrder = " order by m.[member_login]" & sDirection
    if iSort = 2 then sOrder = " order by m.[first_name]" & sDirection
    if iSort = 3 then sOrder = " order by m.[last_name]" & sDirection
    if iSort = 4 then sOrder = " order by m.[member_level]" & sDirection
  end if

'-------------------------------
' Build base SQL statement
'-------------------------------
  sSQL = "select [m].[first_name] as m_first_name, " & _
    "[m].[last_name] as m_last_name, " & _
    "[m].[member_id] as m_member_id, " & _
    "[m].[member_level] as m_member_level, " & _
    "[m].[member_login] as m_member_login " & _
    " from [members] m "
'-------------------------------

'-------------------------------
' Members Open Event begin
' Members Open Event end
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
    SetVar "DListMembers", ""
    Parse "MembersNoRecords", False
    SetVar "MembersNavigator", ""
    Parse "FormMembers", False
    exit sub
  end if
'-------------------------------

'-------------------------------
' Prepare the lists of values
'-------------------------------

  amember_level = Split("1;Member;2;Administrator", ";")
'-------------------------------
'-------------------------------
' Initialize page counter and records per page
'-------------------------------
  iRecordsPerPage = 20
  iCounter = 0
'-------------------------------

'-------------------------------
' Process page scroller
'-------------------------------
  iPage = GetParam("FormMembers_Page")
  if IsEmpty(iPage) then iPage = 1 else iPage = CLng(iPage)
  while not rs.eof and iCounter < (iPage-1)*iRecordsPerPage
    rs.movenext
    iCounter = iCounter + 1
  wend
  iCounter = 0
'-------------------------------

'-------------------------------
' Display grid based on recordset
'-------------------------------
  while not rs.EOF  and iCounter < iRecordsPerPage
'-------------------------------
' Create field variables based on database fields
'-------------------------------
    fldname = GetValue(rs, "m_first_name")
    fldlast_name = GetValue(rs, "m_last_name")
    fldmember_level = GetValue(rs, "m_member_level")
    fldmember_login_URLLink = "MembersInfo.asp"
    fldmember_login_member_id = GetValue(rs, "m_member_id")
    fldmember_login = GetValue(rs, "m_member_login")
'-------------------------------
' Members Show begin
'-------------------------------

'-------------------------------
' Members Show Event begin
' Members Show Event end
'-------------------------------

'-------------------------------
' Replace Template fields with database values
'-------------------------------
    
      SetVar "member_login", ToHTML(fldmember_login)
      SetVar "member_login_URLLink", fldmember_login_URLLink
      SetVar "Prmmember_login_member_id", ToURL(fldmember_login_member_id)
      SetVar "name", ToHTML(fldname)
      SetVar "last_name", ToHTML(fldlast_name)
      fldmember_level = getValFromLOV(fldmember_level, amember_level)
      SetVar "member_level", ToHTML(fldmember_level)
    Parse "DListMembers", True

'-------------------------------
' Members Show end
'-------------------------------

'-------------------------------
' Move to the next record and increase record counter
'-------------------------------
    rs.MoveNext
    iCounter = iCounter + 1
  wend
'-------------------------------


'-------------------------------
' Members Navigation begin
'-------------------------------
  bEof = rs.eof
  if rs.eof and iPage = 1 then
	SetVar "MembersNavigator", ""
  else
    if bEof then
      SetVar "MembersNavigatorLastPage", "_"
    else
      SetVar "NextPage", (iPage + 1)
    end if
    if iPage = 1 then
      SetVar "MembersNavigatorFirstPage", "_"
    else
      SetVar "PrevPage", (iPage - 1)
    end if
    SetVar "MembersCurrentPage", iPage
    Parse "MembersNavigator", False
  end if
'-------------------------------
' Members Navigation end
'-------------------------------

'-------------------------------
' Finish form processing
'-------------------------------
  set rs = nothing
  SetVar "MembersNoRecords", ""
  Parse "FormMembers", False

'-------------------------------
' Members Close Event begin
' Members Close Event end
'-------------------------------
End Sub
'===============================

%>