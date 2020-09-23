Attribute VB_Name = "ModVBPCorrection"
'Not a perfect test but searches for potentially dangerous code in files that may fire without your knowledge
'
'Thanks to Dean Camera for his article that discussed the evil code that made it necessary to add
'DangerousInitialize and DangerousPath. I also extended the detector to check all code
'EG _Initialize method that could access the disk even from the IDE.
'
'thanks to Karl The Lamer for his upload 'Watch Out' txtCodeId=61095 at PSC
'which first inspired me to add this
'
'Thanks to all those uploads with wrong paths in their vbp files
'Added extra code to check for and auto-correct some incorrect file paths in VBP files
'
'NOTE this code is a subset of my Code Fixer program
'some of the support routines (especially those for identifying real words)for its services
'are far more robust than this program requires and could be simplified
'but the over-all time impact is negligable
'let me know if you think any of them are too excessive.
'
'Thanks and copyrights for other coders may be found through out the code
'
'these explore and attempt to repair that most irritating problem with downloads
'not supplying the correct paths in the vbp file, even if they include the files.
'Also contains a test for potentially dangerous code called from Initialize events which may fire before you have a chance to see it.
'v1.1 added single user safety stamp.
'Allows you to mark a code line to be ignored by Safety Net but only applies to your machine
'
'v1.2 ignore Enabled keyword if called from within a _Timer event
'
'v1.3
'cleaned up the code to remove unnecessary stuff
'changed the Ignore Threat message
'Ignore Threat message now included in the warning message to make it easier to use
'do'h forgot about the vbpParent folder. (also improved the logic of the path switcher code)
'v1.4
'improved file detection by using the safer FileSystemObject FileExists
Option Explicit
Public VBInstance                          As VBIDE.VBE
Public strMode                             As String
Private Const READ_CONTROL                 As Long = &H20000
Private Const STANDARD_RIGHTS_READ         As Long = (READ_CONTROL)
Private Const SYNCHRONIZE                  As Long = &H100000
Private Const KEY_ENUMERATE_SUB_KEYS       As Long = &H8
Private Const KEY_NOTIFY                   As Long = &H10
Private Const KEY_QUERY_VALUE              As Long = &H1
Private Const KEY_READ                     As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const HKEY_LOCAL_MACHINE           As Long = &H80000002
Private Const ERROR_NONE                   As Long = 0
Public DangerFound                         As Boolean
'Private bForceShow                         As Boolean
' Let FInd control in code fil list without showing codepane
Public Enum eVBPSearch
  eCorrect
  eVBPFolder
  eSubFolder
  eVBPParent
  eNeighbourFolder
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eCorrect, eVBPFolder, eSubFolder, eVBPParent, eNeighbourFolder
#End If
Public Enum eFileWarningConditions
  eNone
  eSeparateBranch
  eDistant
  eDifferentDisk
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eNone, eSeparateBranch, eDistant, eDifferentDisk
#End If
Public ArrFuncPropSub                      As Variant
Public arrMaliciousStart                   As Variant
Public Const WARNING_MSG                   As String = "'WARNING: "
Private Const SUGGESTION_MSG               As String = "'SUGGESTION: To make Safety Net ignore this code leave the next comment line in code and restore the code" & vbNewLine
Private Const strNotLetterFilter           As String = "[!_!a-z!A-Z!À-Ö!Ø-ß!à-ö!ø-ÿ!0-9]"
Private FSO                                As New FileSystemObject
Private Const DQuote                       As String = """"
Private Const SQuote                       As String = "'"
Private Const SngSpace                     As String = " "
Public StrGuard                            As String
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                            ByVal lpValueName As String, _
                                                                                            ByVal lpReserved As Long, _
                                                                                            lpType As Long, _
                                                                                            ByVal lpData As String, _
                                                                                            lpcbData As Long) As Long


Public Function ArrayMember(ByVal Tval As Variant, _
                            ParamArray pMembers() As Variant) As Boolean

  Dim wal As Variant

'returns true if any member of pMembers equals Tval
  For Each wal In pMembers
    If Tval = wal Then
      ArrayMember = True
      Exit For
    End If
  Next wal

End Function

Private Sub AutoPathFix(strVBPFile As String)

'v4.0.0 search vbp file and check paths are correct,
'Auto repair if file is available in vbp folder or sub-folders
'VB will be able to continue loading

  Dim arrTmp         As Variant
  Dim EPos           As Long
  Dim lHit           As Long
  Dim TStream        As TextStream
  Dim I              As Long
  Dim strBackUpVBP   As String
  Dim StrCorrectPath As String
  Dim BadLoc         As eFileWarningConditions
  Dim StrSubFolder   As String

  Set TStream = FSO.OpenTextFile(strVBPFile, ForReading)
  arrTmp = Split(TStream.ReadAll, vbNewLine)
  For I = LBound(arrTmp) To UBound(arrTmp)
    EPos = InStr(arrTmp(I), "=") - 1
    If EPos > 0 Then
      If ArrayMember(Left$(arrTmp(I), EPos), "Form", "Module", "UserControl", "UserDocument", "Class", "PropertyPage") Then
        If Not VBPPathOK(strVBPFile, arrTmp(I), eCorrect, StrCorrectPath, BadLoc) Then
'test path is correct
          If VBPPathOK(strVBPFile, arrTmp(I), eVBPFolder) Then
'test if file is available in vbp folder
            If InStr(arrTmp(I), ";") = 0 Then
              arrTmp(I) = Left$(arrTmp(I), InStr(arrTmp(I), "=")) & Mid$(arrTmp(I), InStrRev(arrTmp(I), "\") + 1)
             Else
              arrTmp(I) = Left$(arrTmp(I), InStr(arrTmp(I), ";")) & Mid$(arrTmp(I), InStrRev(arrTmp(I), "\") + 1)
            End If
            lHit = lHit + 1
           Else
            If VBPPathOK(strVBPFile, arrTmp(I), eSubFolder, StrSubFolder) Then
'test if file is available in sub-folder
              arrTmp(I) = RewriteVBPLine(arrTmp(I), StrSubFolder, strVBPFile)
              lHit = lHit + 1
             Else
              If VBPPathOK(strVBPFile, arrTmp(I), eVBPParent, StrSubFolder) Then
                arrTmp(I) = RewriteVBPLine(arrTmp(I), StrSubFolder, strVBPFile)
                lHit = lHit + 1
               Else
                If VBPPathOK(strVBPFile, arrTmp(I), eNeighbourFolder, StrSubFolder) Then
                  arrTmp(I) = RewriteVBPLine(arrTmp(I), StrSubFolder, strVBPFile)
                  lHit = lHit + 1
                End If
              End If
            End If
          End If
         Else
          Select Case BadLoc
           Case eDistant
            MsgBox "The file " & strInSQuotes(FileNameOnly(StrCorrectPath)) & " is located 2 or more folders from the vbp folder." & vbNewLine & _
       "While legal this may make the code hard to manage.", vbInformation, "SAFETY NET - Poor File Location"
           Case eSeparateBranch
            MsgBox "The file " & strInSQuotes(FileNameOnly(StrCorrectPath)) & " is located on a different branch of the disk folder heirarchy." & vbNewLine & _
       "While legal this may make the code hard to manage.", vbInformation, "SAFETY NET - Poor File Location"
           Case eDifferentDisk
            MsgBox "The file " & strInSQuotes(FileNameOnly(StrCorrectPath)) & " is not on the same drive as the vbp file." & vbNewLine & _
       "While legal this may make the code hard to manage.", vbInformation, "SAFETY NET -  Poor File Location"
          End Select
        End If
      End If
    End If
  Next I
  If lHit Then
    strBackUpVBP = UCase$(Replace$(LCase$(strVBPFile), ".vbp", "OLD.vbp"))
    FSO.CopyFile strVBPFile, strBackUpVBP 'overwrites if necessary
    Set TStream = FSO.OpenTextFile(strVBPFile, ForWriting)
    TStream.Write (Join(arrTmp, vbNewLine))
    Set TStream = Nothing
    MsgBox "Safety Net has peformed " & lHit & " repair" & IIf(lHit = 1, vbNullString, "s") & " on the VBP file : " & strInSQuotes(FileNameOnly(strVBPFile)) & "." & vbNewLine & _
       vbNullString & vbNewLine & _
       "Files with incorrect paths but available in or near the VBP folder have been given the correct path." & vbNewLine & _
       "Code will now load." & vbNewLine & _
       "The old vbp file has been backed-up to " & strInSQuotes(FileNameOnly(strBackUpVBP)) & "." & vbNewLine & _
       vbNewLine & _
       "To Avoid this problem in future you should do one of 2 things:" & vbNewLine & _
       "1. In VB: Use the 'File|Save<filename> as' menu to move the files into the same folder as the vbp file" & vbNewLine & _
       "OR" & vbNewLine & _
       "2. In File Explorer: Move the files and edit the vbp file in NotePad.", vbInformation, "SAFETY NET - VBP PATH ERROR AUTO-REPAIRED"

  End If

End Sub

Public Function Between(ByVal Lo As Long, _
                        ByVal Tval As Long, _
                        ByVal Hi As Long, _
                        Optional ByVal Exclusive As Boolean = False) As Boolean

  Between = Hi >= Tval And Lo <= Tval
  If Exclusive Then
    Between = Hi > Tval And Lo < Tval
  End If

End Function

Private Function CommentClip(varSearch As Variant) As String

  Dim MyStr       As String
  Dim CommentPos  As Long
  Dim SpaceOffSet As Long

'This code clips end comments from VarSearch
'NOTE also Modifies VarSearch
'UPDATE now copes with literal embedded '
  On Error GoTo BadError
  MyStr = varSearch
  CommentPos = InStr(1, MyStr, SQuote)
  If CommentPos > 0 Then
    Do While InLiteral(MyStr, CommentPos)
      CommentPos = InStr(CommentPos + 1, MyStr, SQuote)
      If CommentPos = 0 Then
        Exit Do
      End If
    Loop
    If CommentPos > 0 Then
      CommentClip = Mid$(MyStr, CommentPos)
      MyStr = Left$(MyStr, CommentPos - 1)
'Preserve spaces with comment if comment is offset with them
      SpaceOffSet = Len(MyStr) - Len(RTrim$(MyStr))
      CommentClip = String$(SpaceOffSet, 32) & CommentClip
      varSearch = Left$(MyStr, Len(MyStr) - SpaceOffSet)
'v 2.4.7 Thanks Evan Toder and Lawrence Miller
' special case of Long2Int where code is of form 'Function Wally(X as Integer) As Integer'
' the X is updated first then a comment with a newline is added
' this stops the Function Type updating and goes into endless loop
      If Len(varSearch) Then
'v2.5.0
        Do While InStr(vbNewLine, Right$(varSearch, 1))
          CommentClip = CommentClip & Right$(varSearch, 1)
          varSearch = Left$(varSearch, Len(varSearch) - 1)
        Loop
      End If
    End If
  End If
  On Error GoTo 0

Exit Function

BadError:
  CommentClip = vbNullString

End Function

Public Function ContainsWholeWord(ByVal strSearch As String, _
                                  ByVal strFind As String, _
                                  Optional ByVal Start As Long = 1, _
                                  Optional CaseSensitive As VbCompareMethod = vbBinaryCompare) As Boolean

  strSearch = Mid$(strSearch, Start)
  If CaseSensitive = vbBinaryCompare Then
    ContainsWholeWord = SngSpace & strSearch & SngSpace Like "*" & strNotLetterFilter & strFind & strNotLetterFilter & "*"
    If Not ContainsWholeWord Then
      ContainsWholeWord = Left$(strSearch, Len(strFind) + 1) = strFind & " " And Left$(strSearch, 1) <> "'"
    End If
   Else
    ContainsWholeWord = LCase$(SngSpace & strSearch & SngSpace) Like "*" & strNotLetterFilter & LCase$(strFind) & strNotLetterFilter & "*"
  End If

End Function

Public Function CountSubString(ByVal varSearch As Variant, _
                               ByVal varFind As Variant) As Long

  CountSubString = UBound(Split(varSearch, varFind))

End Function

Private Sub DangerousInitialize(arrTmp As Variant, _
                                strFile As String)

'CF generated Sub 15/07/2005
'
'Detect any code that could threaten your system by being placed such that simply viewing the
'Designers (usually Forms with UserControls) will trigger an action. The only reason for this
'I can think of is nasty, so this disables suspicious code
'If you have a legitimate reason for this you should accept the cost of having CF repeatedly
'comment it out in return for no one else doing nasties to you.

  Dim DangerFoundLocal  As Boolean
  Dim strDanger         As String
  Dim I                 As Long
  Dim strHead           As String
  Dim strDangerouscode  As String
  Dim strEnabledWarning As String
  Dim strMsgFileName    As String

'Dim strIgnoreME As Long
  strMsgFileName = strInSQuotes(FileNameOnly(strFile))
  For I = LBound(arrTmp) To UBound(arrTmp)
    If LikeArrayInCode(arrTmp(I), "Sub *_Initialize", "Sub *_InitProperties", "Sub *_ReadProperties", "Sub *_Resize", "Sub *_Show", "Sub *_Activate", "Sub AddinInstance_OnConnection", "Sub *_Timer(") Then
      strHead = strInSQuotes(arrTmp(I))
      I = I + 1
      Do
        If Not JustACommentOrBlank(arrTmp(I)) Then
          If EvilCodeDetector(arrTmp(I), strDanger) Then
            If Not SmartLeft(arrTmp(I - 1), StrGuard) Then
              strDangerouscode = strInSQuotes(arrTmp(I))
              If strDanger = "'Enabled'" And arrTmp(I) Like "*Sub *_Timer(*" And strMode = "Soft" Then
'v1.2 ignore timers enabling timers
                GoTo NotDangerous
              End If
              If strDanger = "'Print'" Then
                If InStr(arrTmp(I), "Debug.Print") Then
' the one  safe use of Print
                  GoTo NotDangerous
                End If
              End If
              If strDanger = "'Enabled'" Then
                strEnabledWarning = "Also check the effect of Activating the control that is enabled for other potentially dangerous code."
               Else
                strEnabledWarning = vbNullString
              End If
              MsgBox strHead & " in file " & strMsgFileName & vbNewLine & _
       "contains potentially dangerous code: " & strDanger & " in the code line:" & vbNewLine & _
       strDangerouscode & vbNewLine & _
       IIf(LenB(strEnabledWarning), strEnabledWarning & vbNewLine, "") & "The code has been disabled by commenting it out." & vbNewLine & _
       "Apologies if you have a legitimate reason for this code." & vbNewLine & _
       "For safety reasons there is no way to stop Safety Net doing this each time you open the project." & vbNewLine & _
       "Safety Net will automatically take you to the relevant code when loading is completed.", vbCritical + vbApplicationModal, "SAFETY NET - POTENTIALLY DANGEROUS CODE DETECTED"
              DangerFoundLocal = True
              arrTmp(I) = "'----------------------------------------------" & vbNewLine & _
                          WARNING_MSG & "DANGEROUS CODE DISABLED FOR SAFETY PURPOSES (" & strDanger & ")" & vbNewLine & _
                          IIf(LenB(strEnabledWarning), WARNING_MSG & UCase$(strEnabledWarning) & vbNewLine, vbNullString) & WARNING_MSG & "MAKE SURE YOU KNOW WHAT THIS CODE DOES BEFORE REACTIVATING IT." & vbNewLine & _
                          WARNING_MSG & "CODE IN INITIALIZE PROCEDURES WILL FIRE IF YOU SIMPLY" & vbNewLine & _
                          WARNING_MSG & "VIEW A DESIGNER THAT SHOWS THE FORM OR USER CONTROL IN THE IDE" & vbNewLine & _
                          SUGGESTION_MSG & StrGuard & " " & Now & vbNewLine & _
                          "''" & arrTmp(I) & vbNewLine & _
                          "'----------------------------------------------"
              If strDangerouscode Like "*" & DQuote & "*:*\*" & DQuote & "*" Then
                MsgBox strHead & " in file " & strMsgFileName & vbNewLine & _
       "contains potentially dangerous code referencing a hard-coded path in the code line:" & vbNewLine & _
       strDangerouscode & vbNewLine & _
       "The code has been disabled by commenting it out." & vbNewLine & _
       "Apologies if you have a legitimate reason for this code." & vbNewLine & _
       "For safety reasons there is no way to stop Safety Net doing this each time you open the project." & vbNewLine & _
       "Safety Net will automatically take you to the relevant code when loading is completed.", vbCritical + vbApplicationModal, "SAFETY NET - POTENTIALLY DANGEROUS CODE DETECTED"
                DangerFoundLocal = True
                arrTmp(I) = "'----------------------------------------------" & vbNewLine & _
                            WARNING_MSG & "DANGEROUS CODE DISABLED FOR SAFETY PURPOSES (hard-coded path)" & vbNewLine & _
                            WARNING_MSG & "HARD-CODED PATHS IN INITIALIZATION MAY GIVE CODE ACCESS TO FILES" & vbNewLine & _
                            IIf(LenB(strDanger), vbNullString, WARNING_MSG & " IF YOU SIMPLY VIEW A DESIGNER THAT SHOWS THE FORM OR USER CONTROL IN THE IDE" & vbNewLine) & IIf(LenB(strDanger), vbNullString, SUGGESTION_MSG & StrGuard & " " & Now & vbNewLine) & "''" & arrTmp(I) & vbNewLine & _
                            "'----------------------------------------------"
              End If
            End If
          End If
        End If
NotDangerous:
        I = I + 1
        If I > UBound(arrTmp) Then               '<< HEAVY IDIOT PROOFING
          Exit Do
        End If
      Loop Until Left$(Trim$(CStr(arrTmp(I))), 7) = "End Sub"
      If DangerFoundLocal Then
        FSO.OpenTextFile(strFile, ForWriting).Write Join(arrTmp, vbNewLine)
        DangerFound = True
      End If
    End If
  Next I

End Sub

Private Sub DangerousPath(arrTmp As Variant, _
                          strFile As String)

'CF generated Sub 15/07/2005
'Detects any code that combines potential dangerous commands with explicit full path strings
'It is generally not safe to use hard-coded paths and definitely not in conjunction with code that
'may cause a change to a specific file.
'If you have a legitimate reason for this you should accept the cost of having CF repeatedly
'comment it out in return for no one else doing nasties to you.

  Dim DangerFoundLocal As Boolean
  Dim strDanger        As String
  Dim J                As Long
  Dim strHead          As String
  Dim strMsgFileName   As String
  Dim strDangerouscode As String
  strMsgFileName = strInSQuotes(FileNameOnly(strFile))
  For J = LBound(arrTmp) To UBound(arrTmp)
    If isProcHead(arrTmp(J)) Then
      strHead = strInSQuotes(arrTmp(J))
    End If
    If arrTmp(J) Like "*" & DQuote & "*:*\*.*" & DQuote & "*" Then
      If Not JustACommentOrBlank(arrTmp(J)) Then
        If EvilCodeDetector(arrTmp(J), strDanger) Then
          If Not SmartLeft(arrTmp(J - 1), StrGuard) Then
            strDangerouscode = strInSQuotes(arrTmp(J))
            If strDanger = "Print" Then
              If InStr(arrTmp(J), "Debug.Print") Then
' the one safe use of Print ?
                GoTo NotDangerous
              End If
            End If
            MsgBox strHead & " in file " & strMsgFileName & vbNewLine & _
       "contains potentially dangerous code: " & strDanger & " in association with a hard-coded file path:" & vbNewLine & _
       strDangerouscode & vbNewLine & _
       "The code has been disabled by commenting it out." & vbNewLine & _
       "Apologies if you have a legitimate reason for this code." & vbNewLine & _
       "For safety reasons there is no way to stop Safety Net doing this each time you open the project." & vbNewLine & _
       "Safety Net will automatically take you to the relevant code when loading is completed.", vbCritical, "SAFETY NET - POTENTIALLY DANGEROUS CODE DETECTED"
            DangerFoundLocal = True
            arrTmp(J) = "'----------------------------------------------" & vbNewLine & _
                        WARNING_MSG & "DANGEROUS CODE DISABLED FOR SAFETY PURPOSES (" & strDanger & ")" & vbNewLine & _
                        WARNING_MSG & "MAKE SURE YOU KNOW WHAT THIS CODE DOES BEFORE REACTIVATING IT." & vbNewLine & _
                        WARNING_MSG & "HARD-CODE PATHS ARE GENERALLY NOT GOOD PRACTICE" & vbNewLine & _
                        WARNING_MSG & "SAFETY NET ASSUMES THIS CODE MAY BE AN ATTEMPT TO ATTACK YOUR SYSTEM" & vbNewLine & _
                        SUGGESTION_MSG & StrGuard & " " & Now & vbNewLine & _
                        "''" & arrTmp(J) & vbNewLine & _
                        "'----------------------------------------------"
          End If
        End If
      End If
NotDangerous:
      If DangerFoundLocal Then
        FSO.OpenTextFile(strFile, ForWriting).Write Join(arrTmp, vbNewLine)
        DangerFound = True
      End If
      Exit For
    End If
  Next J

End Sub

Private Function dirArray(ByVal strFolder As String) As Variant

'create an array of subfolders for given folder (modifie from VB help file

  Dim Counter     As Long
  Dim strTmp      As String
  Dim dir_array() As String

  On Error Resume Next
  strTmp = Dir(strFolder, vbDirectory)
  Do Until LenB(strTmp) = 0
    strTmp = Dir()
    If LenB(strTmp) Then
      If strTmp <> "." Then
        If strTmp <> ".." Then
          If GetAttr(strFolder & strTmp) And vbDirectory Then
            ReDim Preserve dir_array(Counter)
            dir_array(Counter) = strTmp
            Counter = Counter + 1
          End If
        End If
      End If
    End If
  Loop
  If Counter Then
    dirArray = dir_array
   Else
    dirArray = Split("")
  End If
  On Error GoTo 0

End Function

Private Function EvilCodeDetector(varCode As Variant, _
                                  strDanger As String) As Boolean

'v40.0 scans for suspicious commands in code lines

  Dim K As Long

  strDanger = vbNullString
  For K = LBound(arrMaliciousStart) To UBound(arrMaliciousStart)
    If InStrWholeWordRX(varCode, arrMaliciousStart(K)) Then
      If InCode(varCode, InStr(varCode, arrMaliciousStart(K))) Then
        strDanger = strInSQuotes(arrMaliciousStart(K))
        EvilCodeDetector = True
        Exit For
      End If
    End If
  Next K

End Function

Public Function ExtractCode(varCode As Variant, _
                            Optional StrCom As String = vbNullString, _
                            Optional strSpace As String = vbNullString) As Boolean

  Dim StrOrig As String

'ver 1.1.51
'extracts code only from soucre code
'optionally returns or dumps the left padding spaces and comments
'if there is no code it returns False and resets Varcode to original content
  StrOrig = varCode
  If Len(varCode) Then
    StrCom = CommentClip(varCode)
    strSpace = SpaceOffsetClip(varCode)
    If Len(varCode) Then
      ExtractCode = True
     Else
      varCode = StrOrig
    End If
  End If

End Function

Public Function FileExtention(ByVal filespec As String) As String

  If LenB(filespec) Then
    FileExtention = FSO.GetExtensionName(filespec)
  End If

End Function

Public Function FileNameOnly(ByVal filespec As String) As String

  If LenB(filespec) Then
    FileNameOnly = FSO.GetFileName(filespec)
  End If

End Function

Public Function FilePathOnly(ByVal filespec As String) As String

  If LenB(filespec) Then
    FilePathOnly = FSO.GetParentFolderName(filespec)
  End If

End Function

Public Function GetRegisteredDetails(Target As String) As String

'*PURPOSE: Get Resistry entry for Registered Owner

  Dim hKey   As Long
  Dim BufLen As Long
  Dim BufStr As String

  If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", 0, KEY_READ, hKey) = ERROR_NONE Then
    BufLen = 255
    BufStr = String$(BufLen, vbNullChar)
    If RegQueryValueExString(hKey, Target, 0, 0, BufStr, BufLen) = ERROR_NONE Then
      If BufLen > 0 Then
        GetRegisteredDetails = Left$(BufStr, BufLen - 1)
      End If
    End If
  End If
  RegCloseKey hKey

End Function

Private Function getVBPPath(ByVal strVBP As String, _
                            ByVal FName As String) As String

'open vbp and scan for line that loads a file
'only used to generate message in MissingFileMsg if file can't be located

  Dim arrTmp As Variant
  Dim I      As Long

  arrTmp = Split(FSO.OpenTextFile(strVBP, ForReading).ReadAll, vbNewLine)
  For I = LBound(arrTmp) To UBound(arrTmp)
    If SmartRight(arrTmp(I), FName) Then
      getVBPPath = arrTmp(I)
      Exit For
    End If
  Next I

End Function

Public Function InCode(ByVal varSearch As Variant, _
                       ByVal TestPos As Long) As Boolean

  If TestPos Then
    If InComment(varSearch, TestPos) Then
      InCode = False
     ElseIf InLiteral(varSearch, TestPos) Then
      InCode = False
     ElseIf InTimeLiteral(varSearch) Then
      InCode = False
     Else
      InCode = True
    End If
  End If

End Function

Public Function InComment(ByVal varSearch As Variant, _
                          ByVal TPos As Long) As Boolean

  Dim Possible As Long
  Dim arrTmp   As Variant
  Dim OPos     As Long
  Dim NPos     As Long
  Dim I        As Long

'v2.0.5 fixed it was not hitting properly
  Possible = InStr(varSearch, SQuote)
  If Possible Then
    Do
      If Possible > TPos Then ' the test point is less than the possible point
        Possible = 0
        Exit Do
      End If
      If TPos > Len(varSearch) Then ' the test point is beyond the len of string
        Possible = 0
        Exit Do
      End If
      If InLiteral(varSearch, Possible, False) Then
        Possible = InStr(Possible + 1, varSearch, SQuote)
      End If
      If Possible = 0 Or Possible < TPos Then
        Exit Do
      End If
    Loop While InLiteral(varSearch, Possible, False) And Possible > 0
    Possible = InStr(varSearch, SQuote) < TPos
    If Possible Then
      arrTmp = Split(varSearch, SQuote)
      For I = LBound(arrTmp) To UBound(arrTmp)
        NPos = NPos + 1 + Len(arrTmp(I))
        If Between(OPos, TPos, NPos) Then
          InComment = Not InLiteral(varSearch, OPos, False)
          Exit For
        End If
        OPos = NPos
        If OPos >= TPos Then
          Exit For
        End If
      Next I
    End If
  End If

End Function

Public Function InLiteral(ByVal varSearch As Variant, _
                          ByVal TPos As Long, _
                          Optional ByVal CommentTest As Boolean = True) As Boolean

  Dim Possible As Long
  Dim arrTest  As Variant
  Dim I        As Long
  Dim OPos     As Long
  Dim NPos     As Long

  Possible = InStr(varSearch, DQuote)
  If Possible Then
    If Possible = TPos Then
      InLiteral = Not InComment(varSearch, TPos)
     Else
      arrTest = Split(varSearch, DQuote)
      For I = LBound(arrTest) To UBound(arrTest)
        NPos = NPos + 1 + Len(arrTest(I))
        If NPos > TPos Then
          If IsOdd(I) Then
            If Between(OPos, TPos, NPos) Then
              If CommentTest Then
                InLiteral = Not InComment(varSearch, TPos)
               Else
' this is only to stop nocomment creating recursive overflow
                InLiteral = True
              End If
              Exit For
            End If
          End If
        End If
        OPos = NPos
        If OPos > TPos Then
          Exit For
        End If
      Next I
    End If
  End If

End Function

Public Function InstrArrayLike(varSearch As Variant, _
                               ParamArray varFind() As Variant) As Long

  Dim VarTmp  As Variant
  Dim strTest As String

  strTest = Join(varSearch)
  For Each VarTmp In varFind
    If strTest Like "*" & VarTmp & "*" Then
      InstrArrayLike = True
      Exit For
    End If
  Next VarTmp

End Function

Public Function InStrWholeWordRX(ByVal strSearch As String, _
                                 ByVal strFind As String, _
                                 Optional Start As Long = 1, _
                                 Optional CaseSensitive As VbCompareMethod = vbBinaryCompare) As Long

  Dim TPos As Long

  If LenB(strSearch) Then
    If LenB(strFind) Then
      If Start > 1 Then
        If Start < Len(strSearch) Then
          strSearch = Mid$(strSearch, Start)
         Else
          GoTo SafeExit
        End If
      End If
      If ContainsWholeWord(strSearch, strFind, 1, CaseSensitive) Then
        If Not CaseSensitive = vbBinaryCompare Then
          strSearch = LCase$(strSearch)
          strFind = LCase$(strFind)
        End If
        TPos = InStr(strSearch, strFind)
'Get inital test point then
        If TPos Then
          Do
            If TPos = 1 Then
              If strSearch = strFind Then
                InStrWholeWordRX = 1
                Exit Do
              End If
              If Mid$(strSearch, Len(strFind) + 1, 1) Like strNotLetterFilter Then
                InStrWholeWordRX = 1
                Exit Do
              End If
             ElseIf Mid$(strSearch, TPos - 1, 1) Like strNotLetterFilter Then
              If TPos + Len(strFind) - 1 = Len(strSearch) Then
                InStrWholeWordRX = TPos
                Exit Do
               ElseIf Mid$(strSearch, TPos + Len(strFind), 1) Like strNotLetterFilter Then
                InStrWholeWordRX = TPos
                Exit Do
              End If
            End If
            TPos = InStr(TPos + 1, strSearch, strFind)
          Loop While TPos
        End If
      End If
      If strFind = strSearch Then
        InStrWholeWordRX = 1
      End If
      If InStrWholeWordRX Then
        InStrWholeWordRX = InStrWholeWordRX + Start - 1
      End If
    End If
  End If
SafeExit:

End Function

Private Function InTimeLiteral(ByVal varSearch As Variant) As Boolean

  Dim P1 As Long
  Dim P2 As Long
  Dim Ps As Long

  If CountSubString(varSearch, "#") > 1 Then
    Ps = InStr(varSearch, "#")
    Do
      Do
        P1 = InStr(Ps, varSearch, "#")
        P2 = InStr(P1 + 1, varSearch, "#")
        Ps = P2
        If Ps = 0 Then
          Exit Do
        End If
      Loop While InLiteral(varSearch, P1)
      If P1 > 0 Then
        If Not InComment(varSearch, P1) Then
          If P2 > P1 Then
            If Not InComment(varSearch, P2) Then
              InTimeLiteral = IsDate(Mid$(varSearch, P1, P2))
            End If
          End If
        End If
      End If
      If Ps = 0 Then
        Exit Do
      End If
    Loop While P1 > 0
  End If

End Function

Public Function IsOdd(ByVal N As Variant) As Boolean

'Here's a efficient IsEven function
'By Sam Hills
'shills@bbll.com
'*If you want an IsOdd function, just omit the Not.
'        IsEven =not -(n And 1)

  IsOdd = -(N And 1)

End Function

Public Function isProcHead(ByVal strCode As String) As Boolean

' protect from detecting comments,End Proc or Exit Proc

  Dim ArrCode As Variant
  Dim I       As Long
  Dim J       As Long

  If ExtractCode(strCode) Then
    If LeftWord(strCode) <> "End" Then
      If LeftWord(strCode) <> "Exit" Then
        ArrCode = Split(strCode)
        If UBound(ArrCode) > 1 Then
          For I = 0 To 2
            For J = 0 To 2
              If ArrCode(I) = ArrFuncPropSub(J) Then
                isProcHead = True
                Exit For
              End If
            Next J
            If isProcHead Then
              Exit For
            End If
          Next I
        End If
      End If
    End If
  End If

End Function

Public Sub JumpTo(ByVal strFindMe As String)

  Dim sline As Long
  Dim sCol  As Long
  Dim eline As Long
  Dim eCol  As Long
  Dim Proj  As VBProject
  Dim Comp  As VBComponent

  If LenB(Trim$(strFindMe)) Then
  End If
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then                'Ignore Non-code Related Documents modules
        If Comp.CodeModule.Find(strFindMe, sline, sCol, eline, eCol, True) Then
          Comp.CodeModule.CodePane.Show
          Comp.CodeModule.CodePane.SetSelection sline, sCol, eline, eCol
          GoTo GOTIT
        End If
      End If
    Next Comp
  Next Proj
GOTIT:

End Sub

Public Function JustACommentOrBlank(ByVal varSearch As Variant) As Boolean

'copright 2003 Roger Gilchrist
'detect comments and empty strings
'TestLineSuspension varSearch
'v2.8.3 speed up with inline testing

  varSearch = Trim$(varSearch)
  If LenB(varSearch) = 0 Then
    JustACommentOrBlank = True
   ElseIf Left$(varSearch, 1) = SQuote Then
    JustACommentOrBlank = True
   ElseIf Left$(varSearch, 4) = "Rem " Then
    JustACommentOrBlank = True
  End If
'  safe_sleep

End Function

Public Function LeftWord(ByVal varChop As Variant) As String

  If LenB(varChop) Then
    LeftWord = Split(varChop)(0)
  End If

End Function

Public Function LikeArrayInCode(varSearch As Variant, _
                                ParamArray varFind() As Variant) As Long

  Dim VarTmp     As Variant
  Dim ArrSubTest As Variant
  Dim I          As Long

  For Each VarTmp In varFind
    ArrSubTest = Split(VarTmp, "*")
    If varSearch Like "*" & VarTmp & "*" Then
      For I = UBound(ArrSubTest) To LBound(ArrSubTest) Step -1
        If Len(ArrSubTest(I)) Then
          If InCode(varSearch, InStr(varSearch, ArrSubTest(I))) Then
            LikeArrayInCode = True
            Exit For
          End If
        End If
      Next I
      If LikeArrayInCode Then
        Exit For
      End If
    End If
  Next VarTmp

End Function

Private Sub MissingFileMsg(ByVal strFName As String, _
                           ByVal strVBPPath As String, _
                           MissingCount As Long)

'messagebox for files that can't be found

  MissingCount = MissingCount + 1
  MsgBox "Safety Net could not resolve an incorrect path to file " & strInSQuotes(FileNameOnly(strFName)) & "." & vbNewLine & _
       "VBP PATH: " & getVBPPath(strVBPPath, strFName) & vbNewLine & _
       vbNewLine & _
       "The file may be:" & vbNewLine & _
       "1. Missing (NOTE code may run without it(with a little editing))" & vbNewLine & _
       "2. More than one sub-folder below the vbp folder " & vbNewLine & _
       "3. In a folder more than one above the vbp folder." & vbNewLine & _
       "Safety Net will open the VBP file in NotePad when loading has been completed.", vbCritical, "SAFETY NET - LOADING ERROR " & UCase$(FileNameOnly(strFName))

End Sub

Private Function RewriteVBPLine(varLine As Variant, _
                                strSubF As String, _
                                strVPath As String) As String

  If InStr(varLine, ";") = 0 Then
    RewriteVBPLine = Left$(varLine, InStr(varLine, "=")) & Replace$(strSubF, FilePathOnly(strVPath) & "\", vbNullString)
   Else
    RewriteVBPLine = Left$(varLine, InStr(varLine, ";")) & Replace$(strSubF, FilePathOnly(strVPath) & "\", vbNullString)
  End If

End Function

Public Function SmartLeft(ByVal varSearch As Variant, _
                          varFind As Variant, _
                          Optional ByVal CaseSensitive As Boolean = True) As Boolean

'This routine was originally designed to test multiple possible left strings
'BUT I also use it as a simple way of testing even a single left string
'without having to separately code the length at every instance

  If Len(varSearch) Then
    If Len(varFind) Then
      If Not CaseSensitive Then
        varSearch = LCase$(varSearch)
        varFind = LCase$(varFind)
      End If
      SmartLeft = InStr(varSearch, varFind) = 1
'SmartLeft = Left$(varSearch, Len(varFind)) = varFind
    End If
  End If

End Function

Public Function SmartRight(ByVal varSearch As Variant, _
                           ByVal varFind As Variant, _
                           Optional ByVal CaseSensitive As Boolean = True) As Boolean

'This routine was originally designed to test multiple possible left strings
'BUT I also use it as a simple way of testing even a single left string
'without having to separately code the length at every instance
'CaseSensitive was added to solve a problem with hand coding of standard VB routines with wrong case

  If Not CaseSensitive Then
    varSearch = LCase$(varSearch)
    varFind = LCase$(varFind)
  End If
  SmartRight = Right$(varSearch, Len(varFind)) = varFind

End Function

Private Function SpaceOffsetClip(VarStr As Variant) As String

  Dim CutPoint As Long

'is not always needed but lets some Fixers to operate properly
' by temporarily removeing any leading blanks
  If Left$(VarStr, 1) = SngSpace Then
    CutPoint = 1
    Do While Mid$(VarStr, CutPoint, 1) = SngSpace
      CutPoint = CutPoint + 1
    Loop
    CutPoint = CutPoint - 1
    SpaceOffsetClip = String$(CutPoint, SngSpace)
    VarStr = Mid$(VarStr, CutPoint + 1)
  End If

End Function

Public Function strInSQuotes(varA As Variant) As String

  strInSQuotes = SQuote & varA & SQuote

End Function

Public Sub VBPHandler(ByVal VBProject As VBIDE.VBProject, _
                      FileNames() As String)

  Dim I            As Long
  Dim arrTmp       As Variant
  Dim strFile      As String
  Dim MissingCount As Long

  If LCase$(FileExtention(FileNames(0))) = "vbp" Then
    AutoPathFix FileNames(0)
  End If
'On Error GoTo Oops
  For I = LBound(FileNames) To UBound(FileNames)
    If ArrayMember(LCase$(FileExtention(FileNames(I))), "ctl", "dob", "frm", "pag", "bas", "dsr") Then
      If VBPPathOK(VBProject.FileName, FileNames(I), eCorrect, strFile) Then
        arrTmp = Split(FSO.OpenTextFile(strFile, ForReading).ReadAll, vbNewLine)
        If InstrArrayLike(arrTmp, "Sub *_Initialize", "Sub *_InitProperties", "Sub *_ReadProperties", "Sub *_Resize", "Sub *_Show", "Sub *_Activate", "Sub AddinInstance_OnConnection", "Sub *_Timer(") Then
'check for potentially dangerous locations
          DangerousInitialize arrTmp, strFile
        End If
        If InstrArrayLike(arrTmp, "*" & DQuote & "*:*\*.*" & DQuote & "*") Then
'check for potentially dangerous hard-paths
          DangerousPath arrTmp, strFile
        End If
       Else
'original urpose check for bad paths in vbp files
        MissingFileMsg FileNames(I), VBProject.FileName, MissingCount
      End If
    End If
ResumeSafe:
  Next I
  If MissingCount Then
    Shell "notepad.exe " & VBProject.FileName, vbNormalFocus
    MsgBox "Safety Net has loaded the VBP file into NotePad." & vbNewLine & _
       "Please close VB without saving anything and edit the paths to the files.", vbCritical, "SAFETY NET - LOADING ERRORS: MISSING FILES OR INCORRECT PATHS"
  End If

Exit Sub

Oops:
  Select Case Err.Number
   Case 5, 53, 76
'5 means that the file name extrator ran out of path to trim, 53 is ordinary no file found file above vbp folder, 76 no sub folder found
    MissingFileMsg FileNames(I), VBProject.FileName, MissingCount
   Case Else
    MsgBox "Error(" & Err.Number & ")" & Err.Description, vbCritical, "SAFETY NET - LOADING ERROR"
  End Select
  Resume ResumeSafe

End Sub

Private Function VBPPathOK(ByVal VBPPath As String, _
                           ByVal varPath As Variant, _
                           ByVal vbpLoc As eVBPSearch, _
                           Optional strFullPath As String, _
                           Optional FWarning As eFileWarningConditions = eNone) As Boolean

'v4.0.0 test vbp file paths
'Loc = eCorrect VBP path correct
'Loc = eVBPFolder Path incorrect file in vbp folder
'Loc = eSubFolder Path incorrect file in sub folder
'Loc = eNeighbourFolder  Path incorrect file in sub folder of VBP home folder's immediate parent folder
'updated to use FSO.FileExists

  Dim strFile As String
  Dim strPath As String
  Dim ArrDir  As Variant
  Dim I       As Long

  FSO.GetAbsolutePathName
  FSO.GetStandardStream
  On Error GoTo Oops
  strPath = FilePathOnly(VBPPath)
  strFile = varPath
  FWarning = eNone
  If InStr(strFile, "=") Then
    strFile = Mid$(strFile, InStr(strFile, "=") + 1)
  End If
  If InStr(strFile, ";") Then ' deal with bas & class
    strFile = Trim$(Mid$(strFile, InStr(strFile, ";") + 1))
  End If
  Select Case vbpLoc
   Case eCorrect 'in path as per vbp
    If InStr(strFile, "..\") Then ' check high paths
      Do While Left$(strFile, 2) = ".."
        strFile = Mid$(strFile, 4)
        strPath = Left$(strPath, InStrRev(strPath, "\") - 1)
      Loop
    End If
    If FSO.FileExists(strFile) Then
      If InStr(strFile, "\") = 0 Then
        strFullPath = strPath & "\" & strFile
       Else
        strFullPath = strFile
      End If
     Else
      strFullPath = strPath & "\" & strFile
    End If
    If CountSubString(strFullPath, ":") = 1 Then
      VBPPathOK = FSO.FileExists(strFullPath)
     Else
      strFullPath = strPath & "\" & strFile
      VBPPathOK = FSO.FileExists(strFullPath)
    End If
'MsgBox VBPPathOK & vbNewLine & strFullPath
    If VBPPathOK Then
      If Abs(CountSubString(strFullPath, "\") - CountSubString(strPath & "\", "\")) > 3 Then
        FWarning = eDistant
      End If
      If Left$(VBPPath, 3) <> Left$(strPath, 3) Then
        FWarning = eDifferentDisk
      End If
      If LenB(VBPPath) < LenB(Mid$(strPath, InStrRev(strPath, "\") + 1)) Then
        FWarning = eSeparateBranch
      End If
    End If
   Case eVBPFolder ' wrong path but available in vbp folder
    If InStr(varPath, "\") Then
      strFile = Mid$(varPath, InStrRev(varPath, "\") + 1)
    End If
    strFullPath = strPath & "\" & strFile
    VBPPathOK = FSO.FileExists(strFullPath)
'---------------------------------------------
   Case eSubFolder 'based on help file
    ArrDir = dirArray(strPath & "\")
    If UBound(ArrDir) > -1 Then
      For I = LBound(ArrDir) To UBound(ArrDir)
' you could add an ignore for the original folder but not really necessary
' as the file is missing from there if you reach here
        If FSO.FileExists(strPath & "\" & ArrDir(I) & "\" & strFile) Then
          strFullPath = strPath & "\" & ArrDir(I) & "\" & strFile
          VBPPathOK = FSO.FileExists(strFullPath)
          Exit For
        End If
      Next I
    End If
'---------------------------------------------
   Case eVBPParent
    strPath = Left$(strPath, InStrRev(strPath, "\")) 'go up one folder
    strFullPath = strPath & strFile
    VBPPathOK = FSO.FileExists(strFullPath)
'---------------------------------------------
   Case eNeighbourFolder '
    strPath = Left$(strPath, InStrRev(strPath, "\")) 'go up one folder
    ArrDir = dirArray(strPath)                       'get array of subfolders
    If Not IsEmpty(ArrDir) Then
      If UBound(ArrDir) Then
        For I = LBound(ArrDir) To UBound(ArrDir)
' it represents a directory.
' you could add an ignore for the original folder but not really necessary
' as the file is missing from there if you reach here
          If FSO.FileExists(strPath & ArrDir(I) & "\" & strFile) Then
            strFullPath = strPath & ArrDir(I) & "\" & strFile
            VBPPathOK = FSO.FileExists(strFullPath)
            Exit For
          End If
        Next I
      End If
    End If
  End Select

Exit Function

Oops:
  MsgBox Err.Description

End Function


':)Code Fixer V4.0.0 (Wednesday, 17 August 2005 07:53:47) 80 + 991 = 1071 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|33332222222222222222222222222222|1112222|2221222|222222222233|111111111111|1222222222220|333333|

