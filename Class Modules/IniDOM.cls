VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniDOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'IniDOM
'======
'
'A set of classes for working with private settings files in standard
'Windows INI format.
'
'XML and XML DOM libraries are a lot more heavyweight than most programs
'require.  There are also many kinds of INI-format files in Windows to
'process, from straight .INI files themselves to VB's .DEP files and
'probably many others.
'
'Many programmers fall into the trap of using the system API entrypoints
'GetPrivateProfileString() and WritePrivateProfileString().  While these
'have their place, one should consider the hazards, primarily unexpected
'Registry Mapping of these calls.  It can also be useful to retain,
'update, and insert comments into INI files as well.  The API calls don't
'handle this very well.
'
'This Registry Mapping is getting to be a worse problem with Windows
'Vista and administrative policies that may both turn Mapping on for
'these API calls yet block registry writes.
'
'
'Some INI files may have non-standard or extended syntax that IniDOM
'can not handle.  IniDOM is really meant for storing and retrieving your
'own program's settings and the values in fairly common Windows INI
'files.
'
'IniDOM does handle quoted Key names and values and the simple line-
'folding of long file paths or registry keys that break with a "\" as the
'last character of a continued line:
'
'   SomeKey = C:\Program Files\Office\Word\
'       AddIns\Settings\3705
'
'
'IniDOM does have some quirks.  One "soft" rule seems to be that multiple
'Sections may occur in a file and that the last one overrides earlier
'occurrences.  IniDOM will merge all of the unique Keys it finds under
'multiple Sections with the same name, keeping the last Value for each
'unique Key.  Generally this has the same effect - and may actually be
'the way INI files are normally processed through other software.
'
'Comments and empty lines are retained.  Other lines that are not
'recognized as either Section headings or Key/Value pairs are parsed as
'a special comment-only entry with Unrecognized = True by the Load method
'of IniDOM and saved back literally by the Save method.
'
'
'########### TESTING INIDOM ###########
'
'IniDOM relies on handling several Collection exceptions internally.
'To test it normally please be sure you have the IDE option "Break on
'unhandled errors" set!
'
'######################################
'
'
'The DOM tree is built from a five-level hierarchy of objects:
'
'   IniDOM      is the parent object, the root of the DOM tree.  It
'               contains members:
'
'               Load method       - Loads the DOM from an ADO Stream.
'                                   The Stream must be open, set to
'                                   Type = adTypeText, loaded, and
'                                   with the appropriate CharSet
'                                   property value.
'               Save method       - Saves the DOM to an ADO Stream.
'                                   The Stream must be open and set to
'                                   Type = adTypeText.  Any existing
'                                   contents will be replaced.
'               Settings property - Used to access the IniSections
'                                   collection. DEFAULT PROPERTY.
'
'   IniSections is a collection of IniSection objects.  It contains
'               members:
'
'               Add method        - Adds a new IniSection to the
'                                   collection.
'               Count property    - Number of IniSection objects in
'                                   the collection.
'               Item property     - Retrieves a reference to a specific
'                                   IniSection object in the collection.
'                                   DEFAULT PROPERTY.
'               Remove method     - Removes the specified IniSection from
'                                   the collection.
'               NewEnum method    - Hidden method used to allow For Each
'                                   iteration over the collection.
'
'   IniSection  is an object that contains members:
'
'               Comment property  - String value of comment.
'               Keys property     - Used to access the IniKeys
'                                   collection. DEFAULT PROPERTY.
'               Name property     - String value of name of the Section,
'                                   if "" it signifies a Section-level
'                                   comment-only node.
'               Unrecognized property
'                                 - If True, this is an unrecognized INI
'                                   file line - value stored in Comment.
'
'   IniKeys     is a collection of IniKey objects.  It contains members:
'
'               Add method        - Adds a new IniKey to the collection.
'               Count property    - Number of IniKey object in the
'                                   collection.
'               Item property     - Retrieves a reference to a specific
'                                   IniKey object in the collection.
'                                   DEFAULT PROPERTY.
'               Remove method     - Removes the specified IniKey from
'                                   the collection.
'               NewEnum method    - Hidden method used to allow For Each
'                                   iteration over the collection.
'
'   IniKey      is an object that contains members:
'
'               Comment property  - String value of comment.
'               Name property     - String value of name of the Section,
'                                   if "" it signifies a Section-level
'                                   comment-only node.
'               QuoteName         - Boolean, True for quoted Name.
'               QuoteValue        - Boolean, True for quoted Value.
'               Unrecognized property
'                                 - If True, this is an unrecognized INI
'                                   file line - value stored in Comment.
'               Value property    - String value, the Key's value.
'                                   DEFAULT PROPERTY.
'
'Each collections' own actual key values are upcased for case-insensitive
'searching.  Comment, blank, and unrecognized line items are given a
'unique hidden collection key value internally.  They can only be
'referenced by the collection ordinal (index), which is 1-based.
'
'There are also exception value Enums for several errors:
'
'       IniSectionsExceptions:
'           issNoSection = &H80047D00
'           issDupSection = &H80047D02
'
'       IniKeysExceptions:
'           iksNoKey = &H80047D10
'           iksDupKey = &H80047D12
'
'       IniKeyExceptions:
'           ikeQuoteName = &H80047D20
'           ikeQuoteValue = &H80047D22
'
'The Default Properties and VB type coercion make it more compact to set
'or retrieve items in the DOM:
'
'       idObject("SciCalc")("layout") = 1
'   is: idObject.Sections.Item("SciCalc").Keys.Items("layout").Value = "1"
'
'       idObject("SciCalc").Add "mode", "hex"
'   is: idObject.Sections.Item("SciCalc").Keys.Add "mode", "hex"
'
'       intLayout = idObject("SciCalc")("layout")
'   is: intLayout = _
'           idObject.Sections.Item("SciCalc").Keys.Items("layout").Value
'
'       intLayout = idObject!SciCalc!layout
'   is: intLayout = _
'           idObject.Sections.Item("SciCalc").Keys.Items("layout").Value
'
'       idObject!Font![Arial Bold] = "12pt"
'   is: idObject.Sections.Item("Font").Keys.Items("Arial Bold").Value = _
'           "12pt"
'

Private Const vbQuote As String = """"

Private m_Sections As IniSections

Public Sub Load(ByVal Stream As ADODB.Stream)
Attribute Load.VB_Description = "Loads the DOM from an open ADO Stream object"
    Dim blnQName As Boolean
    Dim blnQValue As Boolean
    Dim strLine As String
    Dim strCurrSection As String
    Dim strFold As String
    Dim strPair() As String
    Dim strParts() As String
    
    With Stream
        .Position = 0
        Do Until .EOS
            'Append to existing folded entry after converting Tab to Space and
            'then Trim-ing leading and trailing Spaces.
            strLine = strFold & Trim$(Replace$(.ReadText(adReadLine), vbTab, " "))
            Select Case Left$(strLine, 1)
                Case ";", ""
                    If Len(strCurrSection) = 0 Then
                        'Section-level comment.
                        m_Sections.Add , Trim$(Mid$(strLine, 2))
                    Else
                        'Key-level comment.
                        m_Sections.Item(strCurrSection).Keys.Add , , , , Trim$(Mid$(strLine, 2))
                    End If
                    
                Case "["
                    strParts = Split(strLine, ";")
                    strParts(0) = Trim$(strParts(0))
                    If Right$(strParts(0), 1) = "]" Then
                        'Found Section header.
                        strCurrSection = Trim$(Mid$(strParts(0), _
                                                    2, _
                                                    Len(strParts(0)) - 2))
                        On Error Resume Next
                        'Ignore "duplicate Section" errors.
                        If UBound(strParts) = 0 Then
                            m_Sections.Add strCurrSection
                        Else
                            m_Sections.Add strCurrSection, Trim$(strParts(1))
                        End If
                        On Error GoTo 0
                    End If
                    'Else ignore malformed header.
                    
                Case Else
                    'Process quoted Name.
                    If Left$(strLine, 1) = vbQuote Then
                        strParts = Split(strLine, vbQuote, 3)
                        If UBound(strParts) = 2 Then
                            blnQName = True
                            strParts(1) = Replace$(strParts(1), "=", vbVerticalTab)
                            strLine = Join(strParts, "")
                        Else
                            blnQName = False
                        End If
                    Else
                        blnQName = False
                    End If
                    'Process quoted Value.
                    strParts = Split(strLine, vbQuote, 3)
                    If UBound(strParts) = 2 Then
                        blnQValue = True
                        strParts(1) = Replace$(strParts(1), ";", vbFormFeed)
                        strLine = Join(strParts, "")
                    Else
                        blnQValue = False
                        strLine = Join(strParts, vbQuote)
                    End If
                    
                    'Break off comment if any.
                    strParts = Split(strLine, ";", 2)
                    If UBound(strParts) = 0 And Right$(strLine, 1) = "\" And Not blnQValue Then
                        'Handle "folding."
                        strFold = strLine
                    Else
                        strFold = ""
                        strPair = Split(strParts(0), "=", 2)
                        If UBound(strPair) = 1 Then
                            'Valid Key/Value pair.
                            
                            'Note: we're reinserting any equal signs in the Name.
                            strPair(0) = RTrim$(Replace$(strPair(0), vbVerticalTab, "="))
                            'Note: we're reinserting any semicolons in the Value.
                            strPair(1) = Trim$(Replace$(strPair(1), vbFormFeed, ";"))
                            On Error Resume Next
                            'Ignore non-existant Sections.  This occurs if we found a
                            'Key/Value pair prior to the first Section.  Even if there
                            'were "blank" Sections (comments or blank lines before the
                            'first Section header) we'll get a no-hit because of the
                            'way we set the collection Key for them.
                            '
                            'Duplicate Keys will overwrite previous value.
                            If UBound(strParts) = 0 Then
                                m_Sections.Item(strCurrSection).Keys.Add strPair(0), _
                                                                         strPair(1), _
                                                                         blnQName, _
                                                                         blnQValue
                            Else
                                m_Sections.Item(strCurrSection).Keys.Add strPair(0), _
                                                                         strPair(1), _
                                                                         blnQName, _
                                                                         blnQValue, _
                                                                         Trim$(strParts(1))
                            End If
                            If Err.Number <> issNoSection Then
                                'Ignore non-existant Sections: on load it means a Key/Pair
                                'prior to any Section heading, which we'll discard.
                                If Err.Number = iksDupKey Then
                                    'Let "duplicate Key" errors slide.
                                    Err.Clear
                                    On Error GoTo 0
                                    With m_Sections.Item(strCurrSection).Keys.Item(strPair(0))
                                        .Value = strPair(1)
                                        .Comment = strParts(1)
                                    End With
                                Else
                                    If Err.Number <> 0 Then Exit Sub 'Let error bubble up.
                                End If
                            Else
                                Err.Clear
                            End If
                            On Error GoTo 0
                        Else
                            'Unrecognized line in the INI stream.
                            If Len(strCurrSection) = 0 Then
                                'Section level.
                                m_Sections.Add , strLine, True
                            Else
                                'Key level.
                                m_Sections.Item(strCurrSection).Keys.Add , , , , strLine, True
                            End If
                        End If
                    End If
            End Select
        Loop
    End With
End Sub

Public Sub Save(ByVal Stream As ADODB.Stream)
Attribute Save.VB_Description = "Saves the DOM to an open ADO Stream object"
    Dim isSection As IniSection
    Dim ikKey As IniKey
    Dim strName As String
    Dim strValue As String
    
    With Stream
        .Position = 0
        .SetEOS
    End With
    
    For Each isSection In m_Sections
        With isSection
            If Len(.Name) <> 0 Then
                Stream.WriteText "[" & .Name & "]", adWriteChar
            End If
            'Else not a Section header.
            
            'Write unrecognized line or finish with comment or just finish line.
            If .Unrecognized Then
                Stream.WriteText .Comment, adWriteLine
            ElseIf Len(.Comment) > 0 Then
                Stream.WriteText ";" & .Comment, adWriteLine
            Else
                Stream.WriteText "", adWriteLine
            End If
                
            For Each ikKey In isSection.Keys
                With ikKey
                    If Len(.Name) > 0 Then
                        'We have a Key/Value pair.
                        If .QuoteName Then
                            strName = vbQuote & .Name & vbQuote
                        Else
                            strName = .Name
                        End If
                        If .QuoteValue Then
                            strValue = vbQuote & .Value & vbQuote
                        Else
                            strValue = .Value
                        End If
                        Stream.WriteText strName & "=" & strValue, adWriteChar
                    End If
                    'Else just a comment or unrecognized line.
                    
                    'Write unrecognized line or finish with comment or just finish line.
                    If .Unrecognized Then
                        Stream.WriteText .Comment, adWriteLine
                    ElseIf Len(.Comment) > 0 Then
                        Stream.WriteText ";" & .Comment, adWriteLine
                    Else
                        Stream.WriteText "", adWriteLine
                    End If
                End With
            Next
        End With
    Next
    
    Stream.Position = 0
End Sub

Public Property Get Sections() As IniSections
Attribute Sections.VB_Description = "Returns Settings collection"
Attribute Sections.VB_UserMemId = 0
    Set Sections = m_Sections
End Property

Private Sub Class_Initialize()
    Set m_Sections = New IniSections
End Sub

Private Sub Class_Terminate()
    Set m_Sections = Nothing
End Sub
