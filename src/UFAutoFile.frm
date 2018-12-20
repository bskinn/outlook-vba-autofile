VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFAutoFile 
   Caption         =   "Select <TYPE>"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   OleObjectBlob   =   "UFAutoFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFAutoFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Dim rootFld As Folder
Public tgtObj As Object
Public catAddStr As String
Public storeName As String
Public baseFldPath As String
Public LBxFontSize As Double
Public popDepth As Long

' =======================  USERFORM  ========================

Private Sub UserForm_Initialize()
    popDepth = -1
    LBxFontSize = 11#
End Sub

Private Sub UserForm_Activate()

    ' Must have the config values set
    If catAddStr = "" Then
        MsgBox "Category string not set!", vbOKOnly + vbCritical, "Error"
        Unload UFAutoFile
        Exit Sub
    End If
    If baseFldPath = "" Then
        MsgBox "Base folder path not set!", vbOKOnly + vbCritical, "Error"
        Unload UFAutoFile
        Exit Sub
    End If
    If storeName = "" Then
        MsgBox "Email store name not set!", vbOKOnly + vbCritical, "Error"
        Unload UFAutoFile
        Exit Sub
    End If
    If popDepth = -1 Then
        MsgBox "Folder depth-to-populate not set!", vbOKOnly + vbCritical, "Error"
        Unload UFAutoFile
        Exit Sub
    End If
    
    Set rootFld = Application.Session.Folders(storeName)
    Set rootFld = folderNavigate(rootFld, baseFldPath)
    
    ' Drop out if target folder not found
    If rootFld Is Nothing Then
        MsgBox "Target folder not found!", vbOKOnly + vbCritical, "Error"
        Unload UFAutoFile
    End If
    
    ' Populate and configure the ListBox
    LBxDests.Font.Size = LBxFontSize
    populateListBox rootFld, popDepth
    
    ' For the first load, move to the top item
    LBxDests.ListIndex = 0
    
End Sub

' =======================  BUTTONS  ========================

Private Sub BtnCancel_Click()
    Unload UFAutoFile
End Sub

Private Sub BtnFile_Click()
    Dim itOb As Object, dstFld As Folder
    Dim row As Long, catStr As String

    If LBxDests.ListIndex < 0 Then Exit Sub
    
    row = LBxDests.ListIndex
    
    If LBxDests.List(row, 1) = "" Then
        ' No subfoldering needed
        Set dstFld = rootFld.Folders(LBxDests.List(row, 0))
    Else
        ' Must dig to subfolders, then find the destination
        Set dstFld = folderNavigate(rootFld, LBxDests.List(row, 1)) _
                            .Folders(LBxDests.List(row, 0))
    End If
    
'    Set dstFld = rootFld.Folders(LBxDests.List(row, 1)) _
'                            .Folders(LBxDests.List(row, 0))
    
    If tgtObj Is Nothing Then
        For Each itOb In ActiveExplorer.Selection
            doMove itOb, dstFld
        Next itOb
    Else
        doMove tgtObj, dstFld
    End If
    
    Unload UFAutoFile
    
End Sub

' ====================  PRIVATE METHODS  =====================

Private Function folderNavigate(rootFld As Folder, fldPath As String) As Folder
    ' Given a path and a root folder, return the folder at that path
    ' Or, return Nothing it can't be reached/found
    
    Dim workFld As Folder
    Dim folderStr As String, remainder As String
    Dim slashLoc As Long, errNum As Long
    
    ' Initialize
    Set workFld = rootFld
    remainder = fldPath
    
    ' Split by backslashes and navigate
    
    Do While Len(remainder) > 0
        ' Locate the backslash, if any
        slashLoc = InStr(remainder, "\")
        If slashLoc > 0 Then
            ' Must split
            folderStr = Left(remainder, slashLoc - 1)
            remainder = Mid(remainder, slashLoc + 1)
        Else
            ' Just use the remainder
            folderStr = remainder
            remainder = ""
        End If
        
        ' If present, navigate. Otherwise, exit with nothing
        On Error Resume Next
        Set workFld = workFld.Folders(folderStr)
        errNum = Err.Number
        Err.Clear
        On Error GoTo 0
        
        Select Case errNum
        Case 0
            ' All is fine, continue
        Case -2147221233
            Set folderNavigate = Nothing
            Exit Function
        Case Else
            ' Reraise
            Err.Raise errNum
        End Select
        
        ' Folder navigation is completed, loop around
    Loop
    
    Set folderNavigate = workFld

End Function

Private Sub populateListBox(workFld As Folder, popDepthVal As Long, _
                            Optional accumPath As String = "")
    
    ' Iterate across the subfolders of the indicated work folder.
    ' If the pop depth value is zero, tag the folder name and
    ' path into the listbox. Otherwise, recurse deeper.
    
    Dim subFld As Folder
    Dim iter As Long
    Dim workPath As String
    
    For Each subFld In workFld.Folders
        If popDepthVal > 0 Then
            ' Need to dig deeper
            If Len(accumPath) < 1 Then
                ' No accumulated path yet, just use this subfolder
                workPath = subFld.name
            Else
                ' There is an accumulated path; accumulate further
                workPath = accumPath & "\" & subFld.name
            End If
            
            ' Recursive call to keep digging down
            populateListBox subFld, popDepthVal - 1, workPath
            
        Else
            ' At the depth to populate from. Accumulate the folders
            If LBxDests.ListCount = 0 Then
                ' Nothing added yet. Just append the thing.
                LBxDests.AddItem subFld.name
                LBxDests.List(0, 1) = accumPath
            Else
                ' Something there, want to add in sorted fashion
                If subFld.name < LBxDests.List(0) Then
                    ' Folder to add sorts earliest in the list
                    ' So, insert at the beginning
                    LBxDests.AddItem subFld.name, 0
                    LBxDests.List(0, 1) = accumPath
                Else
                    ' Doesn't go at the start. Have to search
                    ' for where it goes. Make sure not to crash
                    ' past the end of the list
                    iter = 0
                    Do Until subFld.name < LBxDests.List(iter) Or _
                                iter = LBxDests.ListCount - 1
                        iter = iter + 1
                    Loop
                    
                    ' Regardless of what triggered the stop condition,
                    ' insert the new thing either before or after the
                    ' item at the current 'iter' position of the list,
                    ' based on how it sorts.
                    If subFld.name > LBxDests.List(iter) Then
                        LBxDests.AddItem subFld.name, iter + 1
                        LBxDests.List(iter + 1, 1) = accumPath
                    Else
                        LBxDests.AddItem subFld.name, iter
                        LBxDests.List(iter, 1) = accumPath
                    End If
                End If
            End If
        End If
    Next subFld
    
End Sub

Private Sub doMove(obj As Object, dst As Folder)
    Dim newCatStr As String
    
    newCatStr = makeCatStr(dst.name)
    
    If Len(obj.Categories) < 1 Then
        obj.Categories = newCatStr
    Else
        If InStr(obj.Categories, newCatStr) = 0 Then
            ' Only add indicated category if not already there
            obj.Categories = obj.Categories & ", " & newCatStr
        End If
    End If
    obj.UnRead = False
    
    If Not obj.Parent.FolderPath = dst.FolderPath Then
        obj.Move dst
    End If
    
End Sub

Private Function makeCatStr(dstName As String) As String
    ' Allow custom per-destination category, rather than constant
    Dim rxCat As New RegExp, mchs As MatchCollection
    
    ' If catAddStr starts with two tildes, treat as regex
    ' Otherwise, treat as constant
    If Left(catAddStr, 2) <> "~~" Then
        makeCatStr = catAddStr
    Else
        With rxCat
            .Global = False
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = Mid(catAddStr, 3)
            Set mchs = .Execute(dstName)
        End With
        
        makeCatStr = mchs(0).Value
    End If
    
End Function
