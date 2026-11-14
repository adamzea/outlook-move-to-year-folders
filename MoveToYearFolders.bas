Attribute VB_Name = "MoveToYearFolders"
Option Explicit

' =============================================================================
'  MoveSelectedToYearFolders
'  Outlook Classic (Windows) macro.
'
'  Moves every selected item into a subfolder named for its year
'  ("2025", "2026", ...) inside the folder the item currently lives in.
'  Year folders are created automatically if they don't exist yet.
'
'  Date used for the year, in order of preference:
'    1. ReceivedTime  (received mail; also set on items in Sent Items)
'    2. SentOn        (fallback for sent items without a received time)
'    3. CreationTime  (last resort, e.g. drafts / reports)
'
'  Items that are already inside a folder named for their year are skipped,
'  so running it twice won't create "2025\2025".
' =============================================================================

Private Const NO_DATE As Date = #1/1/4501#

Public Sub MoveSelectedToYearFolders()
    Dim sel As Outlook.Selection
    Dim snapshot As Collection
    Dim itm As Object
    Dim i As Long
    Dim yr As String
    Dim parentFld As Outlook.MAPIFolder
    Dim targetFld As Outlook.MAPIFolder
    Dim folderCache As Object
    Dim yearCounts As Object
    Dim moved As Long, skipped As Long, failed As Long
    Dim summary As String
    Dim k As Variant

    If Application.ActiveExplorer Is Nothing Then
        MsgBox "Open an Outlook folder and select some emails first.", vbExclamation
        Exit Sub
    End If

    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Then
        MsgBox "No emails are selected.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Move " & sel.Count & " selected item(s) into year subfolders " & _
              "(e.g. ""2025"", ""2026"") of the folder they are in?", _
              vbQuestion + vbOKCancel, "Move to Year Folders") <> vbOK Then
        Exit Sub
    End If

    ' Take a snapshot of the selection first. Moving items while looping over
    ' the live Selection can cause items to be skipped.
    Set snapshot = New Collection
    For i = 1 To sel.Count
        snapshot.Add sel.Item(i)
    Next i

    Set folderCache = CreateObject("Scripting.Dictionary")
    Set yearCounts = CreateObject("Scripting.Dictionary")

    For Each itm In snapshot
        yr = ItemYear(itm)

        Set parentFld = Nothing
        On Error Resume Next
        Set parentFld = itm.Parent
        On Error GoTo 0

        If yr = "" Or parentFld Is Nothing Then
            failed = failed + 1
        ElseIf parentFld.Name = yr Then
            ' Already sitting in its year folder.
            skipped = skipped + 1
        Else
            Set targetFld = GetYearFolder(parentFld, yr, folderCache)
            If targetFld Is Nothing Then
                failed = failed + 1
            Else
                On Error Resume Next
                itm.Move targetFld
                If Err.Number = 0 Then
                    moved = moved + 1
                    yearCounts(yr) = yearCounts(yr) + 1
                Else
                    failed = failed + 1
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If

        DoEvents
    Next itm

    ' ---- Summary ------------------------------------------------------------
    summary = "Moved: " & moved & vbCrLf
    For Each k In yearCounts.Keys
        summary = summary & "    " & k & ":  " & yearCounts(k) & vbCrLf
    Next k
    If skipped > 0 Then summary = summary & "Already in year folder (skipped): " & skipped & vbCrLf
    If failed > 0 Then summary = summary & "Could not move: " & failed & vbCrLf

    MsgBox summary, vbInformation, "Move to Year Folders"
End Sub

' Returns the 4-digit year of the item as a string, or "" if no usable date.
Private Function ItemYear(ByVal itm As Object) As String
    Dim d As Date

    d = TryGetDate(itm, "ReceivedTime")
    If Year(d) >= 4501 Then d = TryGetDate(itm, "SentOn")
    If Year(d) >= 4501 Then d = TryGetDate(itm, "CreationTime")

    If Year(d) >= 4501 Then
        ItemYear = ""
    Else
        ItemYear = CStr(Year(d))
    End If
End Function

' Reads a date property by name; returns NO_DATE if the item doesn't have it.
Private Function TryGetDate(ByVal itm As Object, ByVal propName As String) As Date
    TryGetDate = NO_DATE
    On Error Resume Next
    TryGetDate = CallByName(itm, propName, VbGet)
    On Error GoTo 0
End Function

' Finds (or creates) the subfolder named yr under parentFld. Results are cached
' so each folder is looked up only once per run.
Private Function GetYearFolder(ByVal parentFld As Outlook.MAPIFolder, _
                               ByVal yr As String, _
                               ByVal cache As Object) As Outlook.MAPIFolder
    Dim key As String
    Dim f As Outlook.MAPIFolder

    key = parentFld.EntryID & "|" & yr
    If cache.Exists(key) Then
        Set GetYearFolder = cache(key)
        Exit Function
    End If

    On Error Resume Next
    Set f = parentFld.Folders(yr)          ' existing folder?
    If f Is Nothing Then
        Err.Clear
        Set f = parentFld.Folders.Add(yr)  ' create it (same type as parent)
    End If
    On Error GoTo 0

    If Not f Is Nothing Then cache.Add key, f
    Set GetYearFolder = f
End Function
