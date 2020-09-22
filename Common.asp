<!-- #INCLUDE FILE="adovbs.inc" -->
<!-- #INCLUDE FILE="TemplateObj.asp" -->
<%
'
'    Filename: Common.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'
'===============================
' Database Connection Definition
'-------------------------------
' Book Store Connection begin

Dim cn : Set cn = Server.CreateObject("ADODB.Connection")
'-------------------------------
' Create database connection string, login and password variables
'-------------------------------
Dim strConn, strLogin, strPassword
dim dbName
dbName = server.MapPath("BookStore_MSAccess.mdb")
 'd:\Program Files\CodeCharge\Examples\BookStore\BookStore_MSAccess.mdb
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & dbName & ";Persist Security Info=False"
strLogin = "Admin"
strPassword = ""
'-------------------------------
' Open the connection
'-------------------------------
cn.open strConn, strLogin, strPassword
'-------------------------------
' Book Store Connection end
'-------------------------------
' Create forward only recordset using current database and passed SQL statement
'-------------------------------
sub openrs(rs, sql)
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.CursorLocation = adUseServer
  rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
end sub

'-------------------------------
' Create static only recordset using current database and passed SQL statement
'-------------------------------
sub openStaticRS(rs, sql)
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.CursorLocation = adUseServer
  rs.Open sql, cn, adOpenStatic, adLockReadOnly, adCmdText
end sub
'===============================

'===============================
' Site Initialization
'-------------------------------
' Specify Debug mode (true/false)
Dim bDebug : bDebug = false
'-------------------------------
' Obtain the path where this site is located on the server
'-------------------------------
Dim sAppPath : sAppPath = left(Request("PATH_TRANSLATED"), instrrev(Request("PATH_TRANSLATED"), "\"))
'-------------------------------
' Create Header and Footer Path variables
'-------------------------------
Dim sHeaderFileName : sHeaderFileName = sAppPath & "Header.html"
Dim sFooterFileName : sFooterFileName = sAppPath & "Footer.html"
'===============================

'===============================
' Common functions
'-------------------------------
' Convert non-standard characters to HTML
'-------------------------------
function ToHTML(strValue)
  if IsNull(strValue) then 
    ToHTML = ""
  else
    ToHTML = Server.HTMLEncode(strValue)
  end if
end function

'-------------------------------
' Convert value to URL
'-------------------------------
function ToURL(strValue)
  if IsNull(strValue) then strValue = ""
  ToURL = Server.URLEncode(strValue)
end function

'-------------------------------
' Obtain HTML value of a field
'-------------------------------
function GetValueHTML(rs, strFieldName)
  GetValueHTML = ToHTML(GetValue(rs, strFieldName))
end function

'-------------------------------
' Obtain database field value
'-------------------------------
function GetValue(rs, strFieldName)
  on error resume next
  if rs is nothing then
  	GetValue = ""
  elseif (not rs.EOF) and (strFieldName <> "") then
    res = rs(strFieldName)
    if isnull(res) then 
      res = ""
    end if
    if VarType(res) = vbBoolean then
      if res then res = "1" else res = "0"
    end if
    GetValue = res
  else
    GetValue = ""
  end if
  if bDebug then response.write err.Description
  on error goto 0
end function

'-------------------------------
' Obtain specific URL Parameter from URL string
'-------------------------------
function GetParam(ParamName)
  if Request.QueryString(ParamName).Count > 0 then 
    Param = Request.QueryString(ParamName)
  elseif Request.Form(ParamName).Count > 0 then
    Param = Request.Form(ParamName)
  else 
    Param = ""
  end if
  if Param = "" then
    GetParam = Empty
  else
    GetParam = Param
  end if
end function

'-------------------------------
' Convert value for use with SQL statament
'-------------------------------
Function ToSQL(Value, sType)
  Dim Param : Param = Value
  if Param = "" then
    ToSQL = "Null"
  else
    if sType = "Number" then
      ToSQL = replace(CDbl(Param), ",", ".")
    else
      ToSQL = "'" & Replace(Param, "'", "''") & "'"
    end if
  end if
end function

'-------------------------------
' Lookup field in the database based on provided criteria
' Input: Table (Table), Field Name (fName), criteria (sWhere)
'-------------------------------
function DLookUp(Table, fName, sWhere)
  on error resume next
  Dim Res : Res = cn.execute("select " & fName & " from " & Table & " where " & sWhere).Fields(0).Value
  if IsNull(Res) then Res = ""
  DLookUp = Res
  if bDebug then response.write err.Description
  on error goto 0
end function

'-------------------------------
' Obtain Checkbox value depending on field type
'-------------------------------
function getCheckBoxValue(sVal, CheckedValue, UnCheckedValue, sType)
  if isempty(sVal) then
    if UnCheckedValue = "" then
      getCheckBoxValue = "Null"
    else
      if sType = "Number" then
        getCheckBoxValue = UnCheckedValue
      else
        getCheckBoxValue = "'" & Replace(UnCheckedValue, "'", "''") & "'"
      end if
    end if
  else
    if CheckedValue = "" then
      getCheckBoxValue = "Null"
    else
      if sType = "Number" then
        getCheckBoxValue = CheckedValue
      else
        getCheckBoxValue = "'" & Replace(CheckedValue, "'", "''") & "'"
      end if
    end if
  end if
end function

'-------------------------------
' Obtain lookup value from array containing List Of Values
'-------------------------------
function getValFromLOV(sVal, aArr)
  Dim i
  Dim sRes : sRes = ""
  if (ubound(aArr) mod 2) = 1 then
    for i = 0 to ubound(aArr) step 2
      if cstr(sVal) = cstr(aArr(i)) then sRes = aArr(i+1)
    next
  end if
  getValFromLOV = sRes  
end function

'-------------------------------
' Process Errors
'-------------------------------
function ProcessError()
  if cn.Errors.Count > 0 then
    ProcessError = cn.Errors(0).Description & " (" & cn.Errors(0).Source & ")"
  elseif not (Err.Description = "") then
    ProcessError = Err.Description
  else
    ProcessError = ""
  end if
end Function

'-------------------------------
' Verify user's security level and redirect to login page if needed
'-------------------------------
function CheckSecurity(iLevel)
  if Session("UserID") = "" then
    cn.Close
    Set cn = Nothing
    response.redirect("Login.asp?QueryString=" & toURL(request.serverVariables("QUERY_STRING")) & "&ret_page=" & toURL(request.serverVariables("SCRIPT_NAME")))
  else
    if CLng(Session("UserRights")) < CLng(iLevel) then
      cn.Close
      Set cn = Nothing
      response.redirect("Login.asp?QueryString=" & toURL(request.serverVariables("QUERY_STRING")) & "&ret_page=" & toURL(request.serverVariables("SCRIPT_NAME"))) 
    end if
  End if
end function
'===============================

'===============================
'  GlobalFuncs begin
'  GlobalFuncs end
'===============================
%>