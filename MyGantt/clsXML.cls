VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public xmlDoc As MSXML2.DOMDocument
Private xmlNodelist As MSXML2.IXMLDOMNodeList
Public xmlDbName As String
Public xmlValidateOnParse As Boolean
Private xPE As MSXML2.IXMLDOMParseError
Event xmlError(strError As String)
Private strError As String
Private xmlResolveExternals As Boolean
Private xmlPreserveWhiteSpace As Boolean

Private Function File_Exists(ByVal strFile As String) As Boolean
    On Error Resume Next
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    File_Exists = fs.FileExists(strFile)
    Set fs = Nothing
    Err.Clear
End Function

Public Function Database_Exists(ByVal strDB As String) As Boolean
    Database_Exists = File_Exists(strDB)
End Function

Public Property Get State() As String
    Select Case xmlDoc.readyState
    Case 0
        State = "Uninitialized"
        'loading has not started
    Case 1
        State = "Loading"
        'the load method is executing
    Case 2
        State = "Loaded"
        'load method is complete
    Case 3
        State = "Interactive"
        'enough of the DOM is available for read-only examination and the data has only been partially parsed
    Case 4
        State = "Completed"
        'data is loaded and parsed and available for read/write operations
    End Select
End Property

Public Function Connection_Open(ByVal strFileName As String) As Boolean
    Set xmlDoc = New MSXML2.DOMDocument
    If File_Exists(strFileName) = False Then
        Connection_Open = False
    Else
        DoEvents
        xmlDoc.async = False        ' completely load the document
        xmlDoc.validateOnParse = xmlValidateOnParse
        xmlDoc.resolveExternals = xmlResolveExternals
        xmlDoc.preserveWhiteSpace = xmlPreserveWhiteSpace
        Connection_Open = xmlDoc.Load(strFileName)
        If Connection_Open = False Then
            Set xPE = xmlDoc.parseError
            With xPE
                strError = "Your XML Document failed to load due the following error." & vbCrLf & _
                             "Error #: " & .errorCode & ": " & xPE.reason & _
                            "Line #: " & .Line & vbCrLf & _
                            "Line Position: " & .linepos & vbCrLf & _
                            "Position In File: " & .filepos & vbCrLf & _
                            "Source Text: " & .srcText & vbCrLf & _
                            "Document URL: " & .url
            End With
            RaiseEvent xmlError(strError)
        End If
    End If
    Select Case State
    Case "Completed"
        Connection_Open = True
    Case Else
        Connection_Open = False
    End Select
    xmlDbName = strFileName
End Function

Public Function Connection_Close() As Boolean
    If State = "Completed" Then
        xmlDoc.save xmlDbName
        Set xmlDoc = Nothing
        Connection_Close = True
    Else
        Connection_Close = False
    End If
End Function

Public Function Database_Create(ByVal strDbName As String, Optional Replace As Boolean = False, Optional strComment As String = vbNullString) As Boolean
    Dim strXML As String
    Dim xmlNode As MSXML2.IXMLDOMNode
        
    strXML = "<?xml version='1.0'?>" & vbNewLine
    strXML = strXML + "<tables>" & vbNewLine
    strXML = strXML + "</tables>" & vbNewLine
    
    If File_Exists(strDbName) = True Then
        If Replace = True Then
            File_Update strDbName, strXML, "w"
            If Len(strComment) > 0 Then
                If Connection_Open(strDbName) = True Then
                    Set xmlNode = xmlDoc.createComment(strComment)
                    xmlDoc.appendChild xmlNode
                    Set xmlNode = Nothing
                    Connection_Close
                End If
            End If
        End If
    Else
        File_Update strDbName, strXML, "w"
        If Len(strComment) > 0 Then
            If Connection_Open(strDbName) = True Then
                Set xmlNode = xmlDoc.createComment(strComment)
                xmlDoc.appendChild xmlNode
                Set xmlNode = Nothing
                Connection_Close
            End If
        End If
    End If
    Database_Create = File_Exists(strDbName)
    End Function

Private Sub File_Update(ByVal filName As String, ByVal filLines As String, Optional ByVal WA As String = "write")
    On Error Resume Next
    Dim cDir As String
    Dim oFS As FileSystemObject
    Dim oStream As TextStream
    'create the directory if it does not exist
    cDir = File_Token(filName, "p")
    If Dir_Exists(cDir) = False Then
        Dir_Create cDir
    End If
    WA = LCase$(Left$(WA, 1))
    Set oFS = New FileSystemObject
    Set oStream = oFS.OpenTextFile(filName, IIf(WA = "w", ForWriting, ForAppending), True)
    oStream.Write filLines & vbCrLf
    oStream.Close
    Set oFS = Nothing
    Set oStream = Nothing
    Err.Clear
End Sub

Private Function File_Token(ByVal strFileName As String, Optional ByVal Sretrieve As String = "F", Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    Dim intNum As Long
    Dim sNew As String
    File_Token = strFileName
    Select Case UCase$(Sretrieve)
    Case "D"
        File_Token = Left$(strFileName, 3)
    Case "F"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, intNum + 1)
        End If
    Case "P"
        If InStr(1, strFileName, Delim, vbTextCompare) > 0 Then
            intNum = InStrRev(strFileName, Delim)
            If intNum <> 0 Then
                File_Token = Mid$(strFileName, 1, intNum - 1)
            End If
        Else
            File_Token = ""
        End If
    Case "E"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, intNum + 1)
        End If
    Case "FO"
        sNew = strFileName
        intNum = InStrRev(sNew, Delim)
        If intNum <> 0 Then
            sNew = Mid$(sNew, intNum + 1)
        End If
        intNum = InStrRev(sNew, ".")
        If intNum <> 0 Then
            sNew = Left$(sNew, intNum - 1)
        End If
        File_Token = sNew
    Case "PF"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            File_Token = Left$(strFileName, intNum - 1)
        End If
    End Select
    Err.Clear
End Function


Private Function StrParse(retarray() As String, ByVal strText As String, ByVal Delimiter As String, Optional RedimensionTo As Long = -1) As Long
    On Error Resume Next
    ' the VB split function clone, this starting at 1
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    varArray = Split(strText, Delimiter)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    varA = VarE + 1
    ReDim retarray(varA)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
        Err.Clear
    Next
    If RedimensionTo <> -1 Then ReDim Preserve retarray(RedimensionTo)
    StrParse = UBound(retarray)
    Err.Clear
End Function

Private Sub ArrayTrimItems(varArray() As String)
    On Error Resume Next
    'trim the array items
    Dim uArray As Long
    Dim cArray As Long
    Dim lArray As Long
    uArray = UBound(varArray)
    lArray = LBound(varArray)
    For cArray = lArray To uArray
        varArray(cArray) = Trim$(varArray(cArray))
        Err.Clear
    Next
    Err.Clear
End Sub

Private Function RemDelim(ByVal Dataobj As String, ByVal Delimiter As String) As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    intDataSize = Len(Dataobj)
    intDelimSize = Len(Delimiter)
    strLast = Right$(Dataobj, intDelimSize)
    Select Case strLast
    Case Delimiter
        RemDelim = Left$(Dataobj, (intDataSize - intDelimSize))
    Case Else
        RemDelim = Dataobj
    End Select
    Err.Clear
End Function


Private Function MvRemoveBlanks(ByVal strValue As String, Optional ByVal Delim As String = ",") As String
    On Error Resume Next
    Dim xData() As String
    Dim xTot As Long
    Dim xCnt As Long
    Dim xRslt As String
    xRslt = ""
    StrParse xData, strValue, Delim
    xTot = UBound(xData)
    For xCnt = 1 To xTot
        If Len(Trim$(xData(xCnt))) > 0 Then
            xRslt = xRslt & Trim$(xData(xCnt)) & Delim
        End If
        Err.Clear
    Next
    xRslt = RemDelim(xRslt, Delim)
    MvRemoveBlanks = xRslt
    Err.Clear
End Function

Public Function Record_Exists(ByVal TbName As String, ByVal ItemName As String, ByVal IdxName As String, ByVal IdxValue As String) As Boolean
    Dim xmlNode As MSXML2.IXMLDOMNode
    Record_Exists = False
    TbName = ProperCase$(TbName)
    IdxName = LCase$(IdxName)
    ItemName = LCase$(ItemName)
    
    Set xmlNodelist = xmlDoc.selectNodes("/tables/" & TbName & "/" & ItemName & "/" & IdxName)
    For Each xmlNode In xmlNodelist
        If LCase(xmlNode.nodeName) = LCase$(IdxName) Then
            If LCase$(xmlNode.Text) = LCase$(IdxValue) Then
                Record_Exists = True
                Exit For
            End If
        End If
    Next
    
    Set xmlNode = Nothing
End Function

Public Function Record_ExistsNew(ByVal TbName As String, ByVal ItemName As String, ByVal IdxName As String, ByVal IdxValue As String) As Boolean
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim curRec As String
    Record_ExistsNew = False
    TbName = ProperCase$(TbName)
    IdxName = LCase$(IdxName)
    ItemName = LCase$(ItemName)
    
    Set xmlNode = xmlDoc.selectSingleNode("/tables/" & TbName & "/" & ItemName & "[" & IdxName & "='" & IdxValue & "']")
    curRec = Replace$(xmlNode.Text, vbTab, "")
    curRec = Replace$(curRec, vbNewLine, "")
    curRec = LCase$(Trim$(curRec))
    Debug.Print curRec, IdxValue
    If curRec = LCase$(IdxValue) Then
        Record_ExistsNew = True
    Else
        Record_ExistsNew = False
    End If
    Set xmlNode = Nothing
End Function

Public Sub Record_Update(ByVal TbName As String, ByVal ItemName As String, ByVal IdxName As String, ByVal IdxValue As String)
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim newRec As MSXML2.IXMLDOMNode
    Dim idxNode As MSXML2.IXMLDOMNode
    Dim itmNode As MSXML2.IXMLDOMNode
    Dim xmlTables As MSXML2.IXMLDOMNode
    Dim recFound As Integer
    Dim recNode As MSXML2.IXMLDOMNode
    TbName = ProperCase$(TbName)
    ItemName = LCase$(ItemName)
    IdxName = LCase$(IdxName)
    
    ' get all indexes
    Set xmlNodelist = xmlDoc.selectNodes("/tables/" & TbName & "/" & ItemName & "/" & IdxName)
    recFound = 0
    For Each xmlNode In xmlNodelist
        If LCase(xmlNode.nodeName) = LCase$(IdxName) Then
            If LCase$(xmlNode.Text) = LCase$(IdxValue) Then
                recFound = 1
                Set newRec = xmlNode
                Exit For
            End If
        End If
    Next
    Select Case recFound
    Case 0
        ' record not found
        Set xmlTables = xmlDoc.selectSingleNode("/tables/" & TbName)
    
        Set idxNode = xmlDoc.createElement(IdxName)
        idxNode.Text = IdxValue
               
        Set itmNode = xmlDoc.createElement(ItemName)
        itmNode.appendChild xmlDoc.createTextNode(vbNewLine + vbTab + vbTab + vbTab)
        itmNode.appendChild idxNode
        itmNode.appendChild xmlDoc.createTextNode(vbNewLine + vbTab + vbTab)
    
        xmlTables.appendChild xmlDoc.createTextNode(vbNewLine + vbTab + vbTab)
        xmlTables.appendChild itmNode
        xmlTables.appendChild xmlDoc.createTextNode(vbNewLine + vbTab)
    Case 1
        ' record found
        'If InStr(1, xmlNode.parentNode.xml, "<holidayyear>", vbTextCompare) = 0 Then
        '    Set idxNode = xmlDoc.createElement("holidayyear")
        '    xmlNode.parentNode.appendChild xmlDoc.createTextNode(vbNewLine + vbTab + vbTab + vbTab)
        '    xmlNode.parentNode.appendChild idxNode
        '    xmlNode.parentNode.appendChild xmlDoc.createTextNode(vbNewLine + vbTab + vbTab)
        'End If
    End Select
    xmlDoc.save xmlDbName
    Set xmlNode = Nothing
End Sub

Public Function Table_Create(ByVal TbName As String, ReplaceIfExists As Boolean, ParamArray FieldNames()) As Boolean
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim spNames() As String
    Dim newTb As MSXML2.IXMLDOMNode
    Dim newFld() As MSXML2.IXMLDOMNode
    Dim xmlTables As MSXML2.IXMLDOMNode
    Dim FieldName As Variant
    Dim strFieldNames As String
    
    TbName = ProperCase(TbName)
    ' put code here to check the table name
    If Table_Exists(TbName) = True Then
        If ReplaceIfExists = True Then
            Table_Delete TbName
        Else
            Table_Create = True
            Exit Function
        End If
    End If
    
    ' build the field names
    strFieldNames = ""
    For Each FieldName In FieldNames
        strFieldNames = strFieldNames & CStr(FieldName) & ","
    Next
    strFieldNames = RemDelim(strFieldNames, ",")
    ' remove blanks
    strFieldNames = MvRemoveBlanks(strFieldNames, ",")
    rsTot = StrParse(spNames, strFieldNames, ",")
    ReDim newFld(rsTot)
    For rsCnt = 1 To rsTot
        rsStr = LCase$(spNames(rsCnt))
        Set newFld(rsCnt) = xmlDoc.createElement(rsStr)
        newFld(rsCnt).Text = ""
        newFld(rsCnt).appendChild xmlDoc.createTextNode(vbNewLine + vbTab + vbTab)
    Next
    ' create the table element
    Set newTb = xmlDoc.createElement(TbName)
    ' append the field names
    newTb.appendChild xmlDoc.createTextNode(vbNewLine + vbTab + vbTab)
    For rsCnt = 1 To rsTot
        newTb.appendChild newFld(rsCnt)
        If rsCnt = rsTot Then
            newTb.appendChild xmlDoc.createTextNode(vbNewLine + vbTab)
        Else
            newTb.appendChild xmlDoc.createTextNode(vbNewLine + vbTab + vbTab)
        End If
    Next
    ' append the table to the tables
    Set xmlTables = xmlDoc.selectSingleNode("tables")
    xmlTables.appendChild xmlDoc.createTextNode(vbNewLine + vbTab)
    xmlTables.appendChild newTb
    xmlTables.appendChild xmlDoc.createTextNode(vbNewLine)
    xmlDoc.save xmlDbName
    Table_Create = Table_Exists(TbName)
End Function

Private Sub Class_Initialize()
    ' by default when loaded, the xml document will be validated against a specified template
    ' turning this off on production is not recommended
    xmlValidateOnParse = False
    xmlResolveExternals = False
    xmlPreserveWhiteSpace = True
End Sub

Public Function Table_Exists(ByVal TbName As String) As Boolean
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim rsStr As String
    
    Table_Exists = False
    TbName = LCase$(TbName)
    Set xmlNodelist = xmlDoc.getElementsByTagName("tables/*")
    For Each xmlNode In xmlNodelist
        rsStr = LCase$(xmlNode.nodeName)
        If rsStr = TbName Then
            Table_Exists = True
            Exit For
        End If
    Next
    Set xmlNode = Nothing
End Function

Public Function Table_Names() As String
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim rsStr As String
    
    rsStr = ""
    Set xmlNodelist = xmlDoc.getElementsByTagName("tables/*")
    For Each xmlNode In xmlNodelist
        rsStr = rsStr & xmlNode.nodeName & ","
    Next
    Table_Names = RemDelim(rsStr, ",")
    Set xmlNode = Nothing
End Function

Public Function Table_FieldNames(ByVal strTbName As String) As String
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim xmlNodeF As MSXML2.IXMLDOMNode
    Dim rsStr As String
    Dim xmlNodelistF As MSXML2.IXMLDOMNodeList

    ' this is case sensitive for table names thus Holidays is not the same as holidays
    rsStr = ""
    Set xmlNodelist = xmlDoc.getElementsByTagName("tables/*")
    For Each xmlNode In xmlNodelist
        If LCase$(strTbName) = LCase$(xmlNode.nodeName) Then
            Set xmlNodelistF = xmlDoc.getElementsByTagName("tables/" & xmlNode.nodeName & "/*")
            For Each xmlNodeF In xmlNodelistF
                rsStr = rsStr & xmlNodeF.nodeName & ","
            Next
            Exit For
        End If
    Next
    
    ' you can also use this code as a replacement of the above
    'rsStr = ""
    'strTbName = ProperCase(strTbName)
    'Set xmlNodelist = xmlDoc.getElementsByTagName("tables/" & strTbName & "/*")
    'For Each xmlNode In xmlNodelist
    '    rsStr = rsStr & xmlNode.nodeName & ","
    'Next
    
    Table_FieldNames = RemDelim(rsStr, ",")
    Set xmlNode = Nothing
    Set xmlNodeF = Nothing
    Set xmlNodelistF = Nothing
End Function

Public Function Table_Delete(ByVal TbName As String) As Boolean
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim rsStr As String
    Dim xmlRoot As MSXML2.IXMLDOMElement
    
    Set xmlRoot = xmlDoc.documentElement

    Table_Delete = False
    TbName = LCase$(TbName)
    Set xmlNodelist = xmlDoc.getElementsByTagName("tables/*")
    For Each xmlNode In xmlNodelist
        rsStr = LCase$(xmlNode.nodeName)
        If rsStr = TbName Then
            xmlRoot.removeChild xmlNode
            xmlDoc.save xmlDbName
            Exit For
        End If
    Next
    Set xmlRoot = Nothing
    Set xmlNode = Nothing
End Function


Private Function Dir_Exists(ByVal strFile As String) As Boolean
    On Error Resume Next
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    Dir_Exists = fs.FolderExists(strFile)
    Set fs = Nothing
    Err.Clear
End Function

Private Sub Dir_Create(ByVal StrCompletePath As String)
    On Error Resume Next
    Dim spPaths() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim curPath As String
    StrCompletePath = File_Validate(StrCompletePath)
    spTot = StrParse(spPaths, StrCompletePath, "\")
    For spCnt = 1 To spTot
        curPath = MvFromMv(StrCompletePath, 1, spCnt, "\")
        If Dir_Exists(curPath) = False Then
            MkDir curPath
        End If
        Err.Clear
    Next
    Err.Clear
End Sub

Private Function File_Validate(ByVal strValue As String) As String
    On Error Resume Next
    Dim fFileN As String
    fFileN = strValue
    fFileN = Replace$(fFileN, "\", "")
    fFileN = Replace$(fFileN, "/", "")
    fFileN = Replace$(fFileN, ":", "")
    fFileN = Replace$(fFileN, "*", "")
    fFileN = Replace$(fFileN, "?", "")
    fFileN = Replace$(fFileN, Chr$(34), "")
    fFileN = Replace$(fFileN, "<", "")
    fFileN = Replace$(fFileN, ">", "")
    fFileN = Replace$(fFileN, "|", "")
    fFileN = Replace$(fFileN, vbTab, "")
    File_Validate = fFileN
    Err.Clear
End Function

Private Function MvFromMv(ByVal strOriginalMv As String, ByVal startPos As Long, Optional ByVal NumOfItems As Long = -1, Optional ByVal Delim As String = vbNullString) As String
    On Error Resume Next
    Dim sporiginal() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sLine As String
    Dim endPos As Long
    sLine = vbNullString
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    StrParse sporiginal, strOriginalMv, Delim
    spTot = UBound(sporiginal)
    If NumOfItems = -1 Then
        endPos = spTot
    ElseIf NumOfItems = -2 Then
        endPos = spTot - 1
    Else
        endPos = (startPos + NumOfItems) - 1
    End If
    For spCnt = startPos To endPos
        If spCnt = endPos Then
            sLine = sLine & sporiginal(spCnt)
        Else
            sLine = sLine & sporiginal(spCnt) & Delim
        End If
        Err.Clear
    Next
    MvFromMv = sLine
    Err.Clear
End Function

Public Function MvField(ByVal strData As String, ByVal fldPos As Long, ByVal Delim As String) As String
    On Error Resume Next
    ' returns a substring from a delimted string
    Dim spData() As String
    Dim spCnt As Long
    MvField = vbNullString
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    Call StrParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case fldPos
    Case -1
        MvField = Trim$(spData(spCnt))
    Case -2
        MvField = Trim$(spData(spCnt - 1))
    Case Else
        If fldPos <= spCnt Then
            MvField = Trim$(spData(fldPos))
        End If
    End Select
    Err.Clear
End Function

Private Function ProperCase(ByVal StrString As String) As String
    On Error Resume Next
    ProperCase = StrConv(StrString, vbProperCase)
    Err.Clear
End Function

