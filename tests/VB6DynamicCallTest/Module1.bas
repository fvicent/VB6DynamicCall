Attribute VB_Name = "Module1"
Option Explicit

Private Declare Sub AttachByte Lib "VB6DynamicCall.dll" Alias "_AttachByte@4" (ByVal value As Byte)
Private Declare Sub AttachInteger Lib "VB6DynamicCall.dll" Alias "_AttachInteger@4" (ByVal value As Integer)
Private Declare Sub AttachDouble Lib "VB6DynamicCall.dll" Alias "_AttachDouble@8" (ByVal value As Double)
Private Declare Sub AttachSingle Lib "VB6DynamicCall.dll" Alias "_AttachSingle@4" (ByVal value As Single)
Private Declare Sub AttachString Lib "VB6DynamicCall.dll" Alias "_AttachString@4" (ByVal value As String)
Private Declare Sub CallByAddress Lib "VB6DynamicCall.dll" Alias "_CallByAddress@4" (ByVal dwAddress As Long)

Public Sub AddBytes(ByVal a As Byte, ByVal b As Byte)
    MsgBox CStr(a) & " + " & CStr(b) & " = " & CStr(a + b)
End Sub

Public Sub AddIntegers(ByVal a As Integer, ByVal b As Integer)
    MsgBox CStr(a) & " + " & CStr(b) & " = " & CStr(a + b)
End Sub

Public Sub AddDoubles(ByVal a As Double, ByVal b As Double)
    MsgBox CStr(a) & " + " & CStr(b) & " = " & CStr(a + b)
End Sub

Public Sub AddSingles(ByVal a As Single, ByVal b As Single)
    MsgBox CStr(a) & " + " & CStr(b) & " = " & CStr(a + b)
End Sub

Public Sub AddStrings(ByVal a As String, ByVal b As String)
    MsgBox a & b
End Sub

Public Sub Main()
    Call AttachByte(8)
    Call AttachByte(6)
    Call CallByAddress(AddressOf AddBytes)
    
    Call AttachInteger(32760)
    Call AttachInteger(3)
    Call CallByAddress(AddressOf AddIntegers)

    Call AttachDouble(1.12)
    Call AttachDouble(0.4)
    Call CallByAddress(AddressOf AddDoubles)

    Call AttachSingle(1.5)
    Call AttachSingle(2.9)
    Call CallByAddress(AddressOf AddSingles)
    
    Call AttachString("Hello ")
    Call AttachString("world")
    Call CallByAddress(AddressOf AddStrings)
End Sub
