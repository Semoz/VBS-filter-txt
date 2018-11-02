Call Write2file("=====new==file====")

Set fs = CreateObject("Scripting.FileSystemObject") 
Set file = fs.OpenTextFile("file.txt", 1, false) 
Do While file.AtEndOfStream <> True
line = file.ReadLine
Call ProcessStr(line)
Loop
file.close 
set fs=Nothing

Sub ProcessStr(txt)
If ubound(split(txt," ")) = 0 Then
content=txt
For i = 1 To Len(content)
   content=ReplaceTest(content,"[^\d\.]","")
Next
If ubound(split(content,".")) >0 Then
   MsgBox(content)
   content=ReplaceTest(content,"\.","")
   MsgBox(content)
End If
if len(content)=6 Then
Call Write2file(txt)
End If
End If
End Sub

Sub Write2file(content)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("out.txt", 8, true)
f.WriteLine(content)
f.Close()
set f = nothing
set fso = Nothing
End Sub

Function RegExpTest(patrn, strng)
Dim regEx, retVal
Set regEx = New RegExp
regEx.Pattern = patrn
regEx.IgnoreCase = False
retVal = regEx.Test(strng)
If retVal Then
    RegExpTest = true
Else
    RegExpTest = false
End If

End Function
Function ReplaceTest(str1, patrn, replStr) 
Dim regEx  
Set regEx = New RegExp
regEx.Pattern = patrn
regEx.IgnoreCase = True
ReplaceTest = regEx.Replace(str1, replStr)
End Function
