Attribute VB_Name = "SetCountry"
Option Explicit

Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
 
' ---------------------------
' Excel 64 bits or 32 bits
' ---------------------------
Dim IsWin64 As Boolean                                ' To know if Windows 64 bits
Dim IsExcel64 As Boolean                              ' To know if Excel 64 bits
' ---------------------------

' ---------------------------
' DNA IntelliSense
' ---------------------------
Const NameSpaceURI_const = "http://schemas.excel-dna.net/intellisense/1.0"
Const SuffixUdfFileName_const = ".IntelliSense.xml"   ' The end part in XLM file name
'Const EncodingXMLHeader_const = "<?xml version=""1.0"" encoding=""UTF-8""?>"  'XML text to add for XML encoding
Dim DNAFileName As String                             ' The DNA IntelliSense file name
' ---------------------------

' ---------------------------
' The UDF Addin
' ---------------------------
Const PrefixUdfFileName_const = "AddTwo"              'The first part in XLM file name. Exple: AddTwo for AddTwo.IntelliSense.xml
Dim AddInPath As String                               ' Exple:"C:\Users\"UserName"\AppData\Roaming\Microsoft\AddIns\"
Dim AddInPathDefault As String                        ' Get the default AddIn path from Application.UserLibraryPath
Public FunctionName  As String                        ' In case it's not the same prefix, he function name itself
Dim AddinParams As Variant                            ' Collection from the Application.AddIns2's parameters:
                                                      '     - "Title"
                                                      '     - "Name"
                                                      '     - "Path"
                                                      '     - "FullName"
                                                      '     - "Installed"

' ---------------------------

' ---------------------------
' Language
' ---------------------------
Const CountryDefault_const = 1033                     ' Default English-US
Const XMLCodes_const = "LanguageCodes.xml"            ' XML File containing the avalable languages values: codes, Names etc ...
Dim UsedFileName As String                            ' XLM file name used by ExcelDna.IntelliSense64.xll. Exple: AddTwo.IntelliSense.xml
Dim CountryFileName As String                         ' XLM file containing the actual country name AND the IntelliSense messages. Exple: AddTwo.IntelliSense.xml
Dim CountryCodeUDF As String                          ' The UDF country code
Dim CountryCodeExcel As String                        ' The Excel country code
' ---------------------------

' The demo function
' Guess!
Public Function AddTwo(aNumber As Long) As Long
  AddTwo = aNumber + 2
End Function

' Guess again!
Public Function Add2Numbers(Number1 As Long, Number2 As Long) As Long
  Add2Numbers = Number1 + Number2
End Function

' GetAddinParams: Return Add-In found with the WhatToFind string. We look into either in AddIns2's Name or Title
' OnlyFirst = True  => ONLY THE FIRST
' OnlyFirst = False => All the found Add-In are returned
'                      It’s up to you to find the one that suits you in the collection
' Each element collection has:
'     - "Title"
'     - "Name"
'     - "Path"
'     - "FullName"
'     - "Installed"
Function GetAddinParams(WhatToFind, Optional OnlyFirst As Boolean = True) As Variant
Dim Result As New Collection '(1 To 5)  'In 1 the Title. In 2 the name, In 3 the Path, In 4 the Fullname, in 5 Installed, 6 AddInNumber
Dim aItem
Dim Param_coll As New Collection
Dim I
        
    On Error Resume Next
    For I = 1 To Application.AddIns2.Count
      If InStr(UCase(Application.AddIns2(I).Name), UCase(WhatToFind)) > 0 Or InStr(UCase(Application.AddIns2(I).Title), UCase(WhatToFind)) > 0 Then
        Param_coll.Add Application.AddIns2(I).Title, "Title"
        Param_coll.Add Application.AddIns2(I).Name, "Name"
        Param_coll.Add Application.AddIns2(I).Path, "Path"
        Param_coll.Add Application.AddIns2(I).FullName, "FullName"
        Param_coll.Add Application.AddIns2(I).Installed, "Installed"
        Param_coll.Add I, "AddInNumber"
        Result.Add Param_coll, CStr(Result.Count + 1)
        If OnlyFirst = True Then
          Exit For
        End If
      End If
      Set Param_coll = Nothing
      Set Param_coll = New Collection
    Next I
    
    Set GetAddinParams = Result
    Err.Clear
    On Error GoTo 0
    Set Result = Nothing
    Set Param_coll = Nothing
End Function

' SetLang: LANGUAGE SETTING
' Description:  The 1rst TO CALL. Set the language.
'               By default the Excel language will be taken into account
'               So far errors messages are shown in a simple MsgBox. Hope to you to modify to manage them.
' ForceLang:  Let you force which language you want.
'             By default it's = False BUT IT IS NOT A bOolean
'             You can set either the code number or the string of SHORT COUNTRY NAME.
Sub SetLang(Optional ForceLang As Variant = False)
Dim FileUDFstr As String  'Set full UDF file name (path and full name)
Dim NameUDFstr As String  'Set the UDF name
Dim iFile As Variant      'use to open file
Dim FileContent As String 'File's contents
Dim Codestr As String     'Get ForceLang Code from LanguageCodes.xml file
Dim Rep As Variant
Dim Temp As Variant

  On Error GoTo ErrHandler
' ---------------------------
' Check if it's Win 64 and/or Excel 64 bits
' ---------------------------
  #If Win64 Then
    IsWin64 = True
  #Else
    IsWin64 = False
  #End If
  #If VBA7 Then
    IsExcel64 = True
  #Else
    IsExcel64 = False
  #End If
' ---------------------------
  
' ---------------------------
' Set the default var. values
' ---------------------------
'  Set the default AddIns path
  AddInPathDefault = Application.UserLibraryPath
'  Set the var DNAFileName with the right file name
  If IsExcel64 = True Then
    DNAFileName = "ExcelDna.IntelliSense64.xll"
  Else
    DNAFileName = "ExcelDna.IntelliSense.xll"
  End If
' ---------------------------

' ---------------------------
' Try to Get ExcelDna.IntelliSense parameters FROM Application.AddIns2
' ---------------------------
'  Get the collection of Addins corresponding to the DNA file
  Set AddinParams = GetAddinParams(DNAFileName, False)
'  If none founed ...
  If AddinParams.Count = 0 Then
    If Len(Trim(Dir(AddInPathDefault & DNAFileName))) > 0 Then
      AddInPath = AddInPathDefault
    Else
      MsgBox DNAFileName & " is not found in AddIns nor in the default directory." & vbCrLf & "Please check if you want the IntelliSense. " & vbCrLf & "THE UDF WILL STILL WORK WITHOUT IT !!", vbApplicationModal + vbExclamation
      Exit Sub
    End If
'  Else we set the vars
  Else
    AddInPath = AddinParams(1)("Path") & "\"
  End If
' ---------------------------

' --------------------------
' Get the add-in path
' --------------------------
' Switch off the error
  On Error Resume Next
'  Check if registered
  If Application.RegisterXLL(AddinParams(1)("FullName")) = False Then
'  Try to register
    Application.RegisterXLL AddinParams(1)("FullName")  'AddInPath & DNAFileName
    If Err.Number <> 0 Then
      MsgBox AddinParams(1)("FullName") & vbCrLf & " is not found in AddIns nor in the default directory." & vbCrLf & "Please check if you want the IntelliSense. " & vbCrLf & "THE UDF WILL STILL WORK WITHOUT IT !!", vbApplicationModal + vbExclamation
      Exit Sub
    End If
  End If
' --------------------------

' --------------------------
' --------------------------
' LANGUAGE SETTING
' --------------------------
' --------------------------

' --------------------------
' Get the Excel country code
' --------------------------
    CountryCodeExcel = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
' --------------------------

' --------------------------
' Read Country in the actual loaded UDF
' --------------------------
    CountryCodeUDF = ReadCountry
' --------------------------

' --------------------------
'  Check if both CountryCode are the same
'  TODO: A way to force the language whatever Excel or Windows language
' --------------------------Delete
  If CountryCodeUDF = CountryCodeExcel And ForceLang = False Then
    Exit Sub
  Else      'Test ForceLang
' If CountryCodeUDF = "0" MEANS NO XML loaded in ThisWorkbook.CustomXMLParts. So we goes on to load one
' FROM LanguageCodes.xml, get the code name (us-ENG; fr-FR ...)
    If (CountryCodeUDF = "0" Or CountryCodeExcel <> CountryCodeUDF) And ForceLang = False Then
      CountryCodeUDF = CountryCodeExcel
      ForceLang = CountryCodeExcel
      Codestr = FindXMLLang(ForceLang, AddInPath & XMLCodes_const)
    ElseIf Not TypeName(ForceLang) = "Boolean" Then
      If IsNumeric(ForceLang) Then
        Codestr = FindXMLLang(ForceLang, AddInPath & XMLCodes_const)
        CountryCodeUDF = ForceLang
      Else
        Rep = FindXMLLang(ForceLang, AddInPath & XMLCodes_const)
        If Not IsEmpty(Rep) Then
          Codestr = Rep
        End If
      End If
    Else
      CountryCodeUDF = CountryCodeExcel
      Codestr = FindXMLLang(CountryCodeUDF, AddInPath & XMLCodes_const)
    End If
  End If

' --------------------------
  
' Delete ALL this function's XML that could be in ThisWorkbook.CustomXMLParts
' If True => At least 1 has been deleted else none
  Temp = DeleteXMLPartByFunctionName(FunctionName)
  
' Build the XML file name used by ExcelDna.IntelliSense64.xll
  UsedFileName = FunctionName & SuffixUdfFileName_const 'PrefixUdfFileName_const & SuffixUdfFileName_const
    
' Set the country XML file name
' Build the Country XML file name to copy
  CountryFileName = AddInPath & FunctionName & "_"   'AddInPath & PrefixUdfFileName_const & "_"
  CountryFileName = CountryFileName & Codestr & ".IntelliSense.xml"

' Copy THE NEW xml file
    If Len(Trim(Dir(CountryFileName))) > 0 Then
      ' Replace the old file
      FileCopy CountryFileName, AddInPath & UsedFileName
      iFile = FreeFile
      Open AddInPath & UsedFileName For Input As #iFile
      ' Load the new XLL content
      FileContent = Input(LOF(iFile), iFile)
      Close #iFile
      
' TODO:  About encoding problem !!
'      FileContent = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & FileContent
      
      ' Add the new XML file in CustomXMLParts
      Temp = FileContent
      ThisWorkbook.CustomXMLParts.Add Temp 'strFileContent
      
      ' Useless for now: Call the Refresh ExcelDna function
      Rep = ReLoadDNA
'     Rep = RefreshDNA   ' Doesn't work so far
      If Rep = False Then
        MsgBox "Excel DNA is not loaded"
      End If
    End If

Exit Sub
ErrHandler:
  Debug.Print Err.Number, "SetLang", Err.Description, Erl
  Err.Clear
  Resume Next
End Sub

' -------------
' Find the country code or the short name
' -------------
' XMLFile = the FULL PATH AND NAME
' Code = Country code OR Short Name to find
' Return -> If Code is numeric then return the code (String) else return the Short Name (String)
Function FindXMLLang(Code, XMLFile) As Variant
Dim xmlDoc As Object
Dim node As Object
Dim resultat As String
  
' Load XML
  Set xmlDoc = CreateObject("MSXML2.DOMDocument")
  xmlDoc.Load XMLFile

' Find the specific language Node
  If IsNumeric(Code) = False Then
    Set node = xmlDoc.SelectSingleNode("//Langue[NomCourt='" & Code & "']")
    If Not node Is Nothing Then
      FindXMLLang = node.SelectSingleNode("CodeRetour").Text
    Else
      FindXMLLang = vbEmpty
    End If
  Else
    Set node = xmlDoc.SelectSingleNode("//Langue[CodeRetour='" & Code & "']")
' NomCourt means ShortName ;)
    If Not node Is Nothing Then
      FindXMLLang = node.SelectSingleNode("NomCourt").Text
    Else
      FindXMLLang = vbEmpty
    End If
  End If

'     Nettoyage
  Set node = Nothing
  Set xmlDoc = Nothing
End Function


' Read the country code in XLM file
Function ReadCountry() As Variant
  Dim xmlPart As CustomXMLPart
  Dim xmlDoc As Object
  Dim CountryNode As Object
  Dim xmlContent As String
  Dim CountryCode As String
  Dim cxn As CustomXMLNode
  Dim XmlNamespace
  On Error GoTo ErrHandler
  
' Get Excel Country code
  CountryCode = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
  If ThisWorkbook.CustomXMLParts.SelectByNamespace(NameSpaceURI_const).Count = 0 Then
    ReadCountry = "0"
    Set CountryNode = Nothing
    Exit Function
  End If
  Set xmlPart = ThisWorkbook.CustomXMLParts.SelectByNamespace(NameSpaceURI_const).Item(1)
  xmlContent = xmlPart.XML
  Set xmlDoc = CreateObject("MSXML2.DOMDocument")
  
  xmlDoc.LoadXML xmlContent
'  Set CountryNode = xmlDoc.SelectSingleNode("//FunctionInfo/Function")
  ' Looking for <Function> elements in the XML with the specific name
  Set CountryNode = xmlDoc.SelectSingleNode("//CountryInfo/Country")
  ' if the fonction is found with, we get the country code Attribute
  If Not CountryNode Is Nothing Then
    ReadCountry = CountryNode.getAttribute("Code")
  End If
  Set CountryNode = Nothing
Exit Function

ErrHandler:
  Debug.Print Err.Number, "ReadCountry", Err.Description, Erl
  Err.Clear
  Resume Next
End Function

' 2) Delete all the function name. Here "AddTwo"
Function DeleteXMLPartByFunctionName(ByVal FunctionName As String) As Boolean
    Dim xmlPart As CustomXMLPart
    Dim xmlDoc As Object
    Dim funcNode As Object
    Dim xmlContent As String
    Dim CountryNode As Object
    DeleteXMLPartByFunctionName = False
    
' Verify all CustomXMLParts in ThisWorkBook
    For Each xmlPart In ThisWorkbook.CustomXMLParts
        ' Load XML content in xmlContent string
        xmlContent = xmlPart.XML
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
' Load the xmlcontent in xmlDoc object
        xmlDoc.LoadXML xmlContent
       ' set the <Function> Node object. Here <Function> is "AddTwo" = FunctionName.
        Set funcNode = xmlDoc.SelectSingleNode("//Function[@Name='" & FunctionName & "']")
        ' If we find it ... delete it
        If Not funcNode Is Nothing Then
            ' Si la fonction correspond, effacer le XML
            xmlPart.Delete
            DeleteXMLPartByFunctionName = True
        End If
    Next xmlPart

End Function

' Load the new XML file
Function RefreshDNA() As Boolean
  Dim buffer As String * 255
  Dim thesize As Long
  Dim Rep
  Dim idParts
  Dim FunctionName
  Dim serverId

  RefreshDNA = False

  thesize = GetEnvironmentVariable("EXCELDNA_INTELLISENSE_ACTIVE_SERVER", buffer, Len(buffer))
  If thesize > 0 Then
    serverId = Left(buffer, thesize)
    
    ' Extract ID
    idParts = Split(serverId, ",")(1)
    
    If Len(idParts) >= 1 Then
      ' Use unique ID to build the name
      FunctionName = "IntelliSenseServerControl_" & idParts
  
      ' Appeler la fonction cachée pour rafraîchir IntelliSense avec l'argument "REFRESH"
      On Error Resume Next
      Rep = Application.Run(FunctionName, "REFRESH")
'      Rep = Application.Run(FunctionName, "DEACTIVATE")
'      Rep = Application.Run(FunctionName, "ACTIVATE")
      On Error GoTo 0
      If Rep Then
        RefreshDNA = True
      Else
         MsgBox "IntelliSense refresh not done."
      End If
    End If
  End If
End Function

Function ReLoadDNA() As Boolean
    Dim xllPath As String
'    Const xllPath = "C:\Users\Papa\AppData\Roaming\Microsoft\AddIns\ExcelDna.IntelliSense64.xll" ' Default XLL path
    Dim addIn As addIn
    Dim wasLoaded As Boolean
    Dim I
    Dim Compteur
    Dim PosNumber
    
    xllPath = AddInPath
    
    ReLoadDNA = False
    Dim aItem
    On Error Resume Next
'   Look for the Add In Excel list.
    For I = 1 To Application.AddIns2.Count
      If InStr(UCase(Application.AddIns2(I).FullName), UCase(xllPath)) > 0 And InStr(UCase(Application.AddIns2(I).Name), UCase(".xll")) > 0 Then
        Compteur = Compteur + 1
        If Application.AddIns2(I).Installed = True Then
          Application.AddIns2(I).Installed = False
        Else
          Application.AddIns2(I).Installed = True
        End If
        If Compteur = 1 Then
          PosNumber = I
          Exit For
        End If
      End If
    Next I
    If Compteur > 0 Then
      If Compteur = 1 Then
        Application.RegisterXLL xllPath
        Application.AddIns2(PosNumber).Installed = False
        Application.Wait Now + TimeValue("00:00:01")
        Application.AddIns2(PosNumber).Installed = True
        ReLoadDNA = True
      Else
        MsgBox "Probleme: Several Excel-Dna IntelliSense are installed"
        ReLoadDNA = False
        Exit Function
      End If
    End If
    
    Err.Clear
    On Error GoTo 0
End Function
