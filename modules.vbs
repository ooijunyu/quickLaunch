' creates character array from input string
function createCharArray(inputString)
    Dim intCounter, intLen, arrChars()

    intLen = Len(inputString)-1
    redim arrChars(intLen)

    For intCounter = 0 to intLen
        arrChars(intCounter) = Mid(inputString, intCounter +1, 1)
    Next

    createCharArray = arrChars
end function

' Launch link through default app using explorer.exe
function launchLink(myLink)
    Dim oShell: Set oShell = WScript.CreateObject("WScript.Shell")
    oShell.Run "cmd.exe /C explorer " & myLink
    Set oShell = Nothing
end function

' Read from text file into dictionary object
function setProgramList(file)
    Dim mySource: Set mySource = CreateObject("Scripting.Dictionary")
    Dim oFileToRead: Set oFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(file,1)
    Dim strLine, splitString

    do while not oFileToRead.AtEndOfStream
        strLine = oFileToRead.ReadLine()
        splitString = Split(strLine, ",")
        mySource.Add LCase(splitString(0)), splitString(1)
    Loop

    oFileToRead.Close
    Set oFileToRead = Nothing
    
    Set setProgramList = mySource
end function

' Compare 2 numbers
function min(a,b)
    If a < b then min = a : Else min = b
end function

' Get the index of minimum value in an array
function getMinIndex(myArray)
    Dim arrayCounter, minimum, minIndex, oArray

    arrayCounter = 0
    minimum = myArray(0)
    minIndex = 0
    
    For Each oArray In myArray
        If oArray < minimum And Not IsEmpty(oArray) Then
            minimum = oArray
            minIndex = arrayCounter
        End If
        arrayCounter = arrayCounter + 1
    Next

    getMinIndex = minIndex
end function

' Get Levenshtein distance between 2 string
' From Rosetta Code
' https://rosettacode.org/wiki/Levenshtein_distance#VBScript
Function Levenshtein(s1, s2)
    Dim d(), i, j, n1, n2, d1, d2, d3
    n1 = Len(s1) + 1
    n2 = Len(s2) + 1
    ReDim d(n1, n2)
    If n1 = 1 Then
        Levenshtein = n2 - 1
        Exit Function
    End If
    If n2 = 1 Then
        Levenshtein = n1 - 1
        Exit Function
    End If
    For i = 1 To n1
        d(i, 1) = i - 1
    Next
    For j = 1 To n2
        d(1, j) = j - 1
    Next
    For i = 2 To n1
        For j = 2 To n2
            d1 = d(i - 1, j    ) + 1
            d2 = d(i,     j - 1) + 1
            d3 = d(i - 1, j - 1) + Abs(Mid(s1, i - 1, 1) <> Mid(s2, j - 1, 1))
            d(i, j) = Min(d1, Min(d2, d3))
        Next
    Next
    Levenshtein = d(n1, n2)
End Function 'Levenshtein