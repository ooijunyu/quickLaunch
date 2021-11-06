
' This script will behave like Windows Run and launch the link of names saved
' in mySource.txt using the best macthed RegEx result by computing minimum
' Levenhstein distance

Option Explicit

' Work as an import function
Sub includeFile(module)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(module).readAll()
End Sub

Dim scriptDir : scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Dim mySource : Set mysource = CreateObject("Scripting.Dictionary")
Dim oRe, userResponse, charArray, counter, obj
Dim matchedArray: matchedArray = Array(1)
Dim levDistArray: levDistArray = Array(1)

' Import functions from modules.vbs
Call includeFile(scriptDir & "\modules.vbs")

' Initialize the dictionary from mySource.txt
Set mySource = setProgramList(scriptDir & "\mySource.txt")

' Setting Up for RegEx search
Set oRe = New RegExp
oRe.Global = True
oRe.IgnoreCase = True

' Show input dialog
userResponse = LCase(InputBox("Type in the program you want to launch.", "Quick Launch"))

' Trigger exit for empty input
If userResponse = "" Then
    WScript.Quit
End If

' Initialize RegEx pattern, matching wildcard in between all characters
charArray = createCharArray(userResponse)
oRe.Pattern = ".*" & Join(charArray, ".*") & ".*"

' Match patterns with dictionary keys, and save matched patterns to mactchedArray()
counter = 0
matchedArray(0) = ""
For Each obj In mySource.keys
    If oRe.test(obj) Then
        matchedArray(counter) = obj 
        counter = counter + 1
        Redim Preserve matchedArray(counter + 1)
    End If   
Next 

' Check Lev Distance of each match
counter = 0
If matchedArray(0) = "" Then
    ' Trigger error dialog if no match is found
    MsgBox "Command Not Found", vbOKOnly + vbExclamation + vbDefaultButton1, "Error"
Else
    ' Launch the min LevDist link
    For Each obj In matchedArray
        If Not IsEmpty(obj) AND obj <> "" Then
            levDistArray(counter) = Levenshtein(obj, userResponse)
            counter = counter + 1
            Redim Preserve levDistArray(counter + 1)
        End If
    Next
    Call launchLink(mySource(matchedArray(getMinIndex(levDistArray))))
End If

WScript.Quit