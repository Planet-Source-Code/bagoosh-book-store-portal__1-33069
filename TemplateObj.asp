<%
'
'    Filename: TemplateObj.asp
'    Generated with CodeCharge 2.0.4
'    ASP 2.0 & Templates.ccp build 11/30/2001
'         
'    Usage:
'     LoadTemplate server.mappath("/templates/new.html"), "main"
'     SetVar "ID", 2
'     SetVar "Value", "Name"
'     Parse "DynBlock", False 'or True if you want to create a list
'     Parse "main", False
'     PrintVar "main"
'

Dim objFSO

Dim DBlocks
Dim ParsedBlocks

Sub SetBlock(sTplName, sBlockName)
  Dim nName
  if not DBlocks.Exists(sBlockName) then
    DBlocks.Add sBlockName, getBlock(DBlocks(sTplName), sBlockName)
  end if
  DBlocks(sTplName) = replaceBlock(DBlocks(sTplName), sBlockName)

  nName = NextDBlockName(sBlockName)
  while not (nName = "")
    SetBlock sBlockName, nName
    nName = NextDBlockName(sBlockName)
  wend
End Sub

Sub LoadTemplate(sPath, sName)
  Dim nName
  if not isObject(objFSO) then 
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set DBlocks = Server.CreateObject("Scripting.Dictionary")
    Set ParsedBlocks = Server.CreateObject("Scripting.Dictionary")
  end if
  if objFSO.FileExists(sPath) then
    DBlocks.Add sName, objFSO.OpenTextFile(sPath).ReadAll
    nName = NextDBlockName(sName)
    while not (nName = "")
      SetBlock sName, nName
      nName = NextDBlockName(sName)
    wend
  end if
End Sub

Sub UnloadTemplate()
  if isObject(objFSO) then 
    Set objFSO = nothing
    Set DBlocks = nothing
    Set ParsedBlocks = nothing
  end if
End Sub

Function GetVar(sName)
  GetVar = DBlocks(sName)
End Function

Function SetVar(sName, sValue)
  if ParsedBlocks.Exists(sName) then
    ParsedBlocks(sName) = replace(replace(sValue, "{", "&#123;"), "}", "&#125;")
  else
    ParsedBlocks.add sName, replace(replace(sValue, "{", "&#123;"), "}", "&#125;")
  end if
End Function

Function Parse(sTplName, bRepeat)
  if ParsedBlocks.Exists(sTplName) then
    if bRepeat then
      ParsedBlocks(sTplName) = ParsedBlocks(sTplName) & ProceedTpl(DBlocks(sTplName))
    else
      ParsedBlocks(sTplName) = ProceedTpl(DBlocks(sTplName))
    end if
  else 
    ParsedBlocks.add sTplName, ProceedTpl(DBlocks(sTplName))
  end if
End Function

Function PrintVar(sName)
  PrintVar = ParsedBlocks(sName)
End function

Function ProceedTpl(sTpl)
  Dim regEx, sMatch, oMatches, sName, sTTpl

  sTTpl = sTpl
  sMatch = getNextPattern(sTTpl, 1)
  while len(sMatch) > 0
    sName = mid(sMatch, 2, len(sMatch) - 2)
    if ParsedBlocks.Exists(sName) then
      sTTpl = replace(sTTpl, sMatch, ParsedBlocks(sName))
    else
      sTTpl = replace(sTTpl, sMatch, DBlocks(sName))
    end if
    sMatch = getNextPattern(sTTpl, 1)
  wend
  ProceedTpl = sTTpl
End Function

Function getNextPattern(str, begin)
  Dim res, b, e, isOk
  Dim w(5)

  b = instr(begin, str, "{")
  if b > 0 then 
    e = instr(b, str, "}")
    w(1) = instr(b, str, " ")
    w(2) = instr(b, str, ";")
    w(3) = instr(b, str, ":")
    w(4) = instr(b, str, "=")
    w(5) = instr(b, str, "(")
    isOk = true
    For i = 1 to 5
      if w(i) < e and w(i) > b then isOk = false
    Next
    if isOk then
      res = mid(str, b, e - b + 1)
    else 
      res = getNextPattern(str, e)
    end if
  else
    res = ""
  end if
  getNextPattern = res
End Function

Function getBlock(sTemplate, sName)
  Dim BBloc, EBlock, alpha
  
  alpha = len(sName) + 12
  BBlock = instr(sTemplate, "<!--Begin" & sName & "-->")
  EBlock = instr(sTemplate, "<!--End" & sName & "-->")
  if not (BBlock = 0 or EBlock = 0) then
    getBlock = mid(sTemplate, BBlock + alpha, EBlock - BBlock - alpha)
  else
    getBlock = ""
  end if
End Function

Function replaceBlock(sTemplate, sName)
  Dim BBloc, EBlock
  
  BBlock = instr(sTemplate, "<!--Begin" & sName & "-->")
  EBlock = instr(sTemplate, "<!--End" & sName & "-->")
  if not (BBlock = 0 or EBlock = 0) then
    replaceBlock = left(sTemplate, BBlock - 1) & "{" & sName & "}" & right(sTemplate, len(sTemplate) - EBlock - len("<!--End" & sName & "-->") + 1)
  else
    replaceBlock = sTemplate
  end if
end function


Function NextDBlockName(sTemplateName)
  dim BTag, ETag, sName, sTemplate
  sTemplate = DBlocks(sTemplateName)
  
  BTag = instr(sTemplate, "<!--Begin")
  if BTag > 0 then
    ETag = instr(BTag, sTemplate, "-->")
    sName = Mid(sTemplate, BTag + 9, ETag - (BTag + 9))
    if instr(sTemplate, "<!--End" & sName & "-->") > 0 then
      NextDBlockName = sName
    else
      NextDBlockName = ""
    end if
  else
    NextDBlockName = ""
  end if
End function

'Print all Dynamic Variables
Function PrintAll()
  dim aPBlocks
  dim aDBlocks
  dim i, res

  aPBlocks = ParsedBlocks.Items
  aDBlocks = DBlocks.Items
  res = "<table border=1>"
  for i = 1 to UBound(aDBlocks)
    res = res & "<tr><td><pre>" & ToHTML(aDBlocks(i)) & "</pre></td></tr>"
  next
  for i = 1 to UBound(aPBlocks)
    res = res & "<tr><td><pre>" & ToHTML(aPBlocks(i)) & "</pre></td></tr>"
  next
  res = res & "</table>"
  PrintAll = res 
End Function
%>