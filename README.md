# VB6DynamicCall
Dynamic argument passing and function calls for VB6.

# Rationale

Functions are not first-class citizens in VB6. All you have is the `AddressOf` keyword in order to get a pointer to a function (specially useful for interfacing with the Win32 API), but there is no simple way to call a function by address nor to figure out the arguments passed to that function at runtime. While an easy task in modern languages, a simple callback mechanism implementation in VB6 leads to boilerplate code and weird workarounds. VB6DynamicCall aims to overcome these limitations.

# Usage

The `CallByAddress()` function lets you call a VB6 function by passing the result of `AddressOf` as an argument. If the target function requires certain arguments, you can "attach" them before making the call. For example, if this is our target function:

```
Public Sub AddIntegers(ByVal a As Integer, ByVal b As Integer)
    MsgBox CStr(a) & " + " & CStr(b) & " = " & CStr(a + b)
End Sub
```

We can invoke `AddIntegers(7, 5)` by doing:

```
Call AttachInteger(7)
Call AttachInteger(5)
Call CallByAddress(AddressOf AddIntegers)
```

Arguments are attached from left to right.

You will find out more examples and the API declarations in the test VB6 application under `tests/VB6DynamicCallTest`.
