Attribute VB_Name = "test"
Option Explicit

Sub test()
Attribute test.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim mySha1 As New SHA, hash_text As String

hash_text = Cells(1, 2)
'mySha1.text = String(1, "abc")
[b2] = mySha1.sha_string(hash_text)
[j2] = mySha1.sha_string(hash_text, "SHA-256")

End Sub


