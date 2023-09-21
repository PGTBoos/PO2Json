' Check if a command-line argument (PO file) is provided
If WScript.Arguments.Count = 0 Then
    WScript.Echo "Usage: Drop a PO file onto this script to convert it to JSON."
    WScript.Quit
End If

' Get the source PO file path from the command-line argument
poFilePath = WScript.Arguments(0)

' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Calculate the output JSON file path in the same directory
outputPath = objFSO.GetParentFolderName(poFilePath) & "\" & objFSO.GetBaseName(poFilePath) & ".json"

' Open the PO file for reading
Set poFile = objFSO.OpenTextFile(poFilePath, 1, False)

' Create a dictionary to store translation data
Set translations = CreateObject("Scripting.Dictionary")

' Read the PO file line by line
Dim line
Dim msgid
Dim msgstr
Do Until poFile.AtEndOfStream
    line = poFile.ReadLine
    If Left(line, 5) = "msgid" Then
        ' Extract msgid
        msgid = Mid(line, 7)
        Do Until poFile.AtEndOfStream
            line = poFile.ReadLine
            If Left(line, 6) = "msgstr" Then
                ' Extract msgstr
                msgstr = Mid(line, 8)
                ' Store in the dictionary
                translations(msgid) = msgstr
                Exit Do
            ElseIf Left(line, 1) = "" Then
                Exit Do
            End If
        Loop
    End If
Loop

' Close the PO file
poFile.Close

' Convert dictionary to JSON
Dim json
json = "{"
q = chr(34)
For Each key In translations.Keys
json = json +  key  + " : " +  translations(key)  + "," 
'    json = json & """" & Replace(key, """", "\""") & """: """ & Replace(translations(key), """", "\""") & ""","
Next
' Remove the trailing comma and close the JSON object
json = Left(json, Len(json) - 1) & "}"

' Create or overwrite the JSON file in the same directory
Set jsonFile = objFSO.CreateTextFile(outputPath, True)
jsonFile.Write json
jsonFile.Close

' Clean up objects
Set objFSO = Nothing
Set translations = Nothing

' Done
WScript.Echo "Conversion complete. JSON file saved to " & outputPath
