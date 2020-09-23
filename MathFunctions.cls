VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function Sec(ByVal X As Currency) As Currency: On Error Resume Next: Sec = 1 / Cos(X): End Function
Function Cosec(ByVal X As Currency) As Currency: On Error Resume Next: Cosec = 1 / Sin(X): End Function
Function Cotan(ByVal X As Currency) As Currency: On Error Resume Next: Cotan = 1 / Tan(X): End Function
Function Arcsin(ByVal X As Currency) As Currency: On Error Resume Next: Arcsin = Atn(X / Sqr(-X * X + 1)): End Function
Function Arccos(ByVal X As Currency) As Currency: On Error Resume Next: Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1): End Function
Function Arcsec(ByVal X As Currency) As Currency: On Error Resume Next: Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1)): End Function
Function Arccosec(ByVal X As Currency) As Currency: On Error Resume Next: Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1)): End Function
Function Arccotan(ByVal X As Currency) As Currency: On Error Resume Next: Arccotan = Atn(X) + 2 * Atn(1): End Function
Function HSin(ByVal X As Currency) As Currency: On Error Resume Next: HSin = (Exp(X) - Exp(-X)) / 2: End Function
Function HCos(ByVal X As Currency) As Currency: On Error Resume Next: HCos = (Exp(X) + Exp(-X)) / 2: End Function
Function HTan(ByVal X As Currency) As Currency: On Error Resume Next: HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X)): End Function
Function HSec(ByVal X As Currency) As Currency: On Error Resume Next: HSec = 2 / (Exp(X) + Exp(-X)): End Function
Function HCosec(ByVal X As Currency) As Currency: On Error Resume Next: HCosec = 2 / (Exp(X) - Exp(-X)): End Function
Function HCotan(ByVal X As Currency) As Currency: On Error Resume Next: HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X)): End Function
Function HArcsin(ByVal X As Currency) As Currency: On Error Resume Next: HArcsin = Log(X + Sqr(X * X + 1)): End Function
Function HArccos(ByVal X As Currency) As Currency: On Error Resume Next: HArccos = Log(X + Sqr(X * X - 1)): End Function
Function HArctan(ByVal X As Currency) As Currency: On Error Resume Next: HArctan = Log((1 + X) / (1 - X)) / 2: End Function
Function HArcsec(ByVal X As Currency) As Currency: On Error Resume Next: HArcsec = Log((Sqr(-X * X + 1) + 1) / X): End Function
Function HArccosec(ByVal X As Currency) As Currency: On Error Resume Next: HArccosec = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X): End Function
Function HArccotan(ByVal X As Currency) As Currency: On Error Resume Next: HArccotan = Log((X + 1) / (X - 1)) / 2: End Function
Function Ð() As Currency: Ð = 3.14159265358979: End Function
Function Pi() As Currency: Pi = 3.14159265358979: End Function
Function Drad() As Currency: Drad = 1.74532925199433E-02: End Function
Function Rdeg() As Currency: Rdeg = 57.2957795130823: End Function
Function LogN(ByVal B As Currency, ByVal r As Currency) As Currency
     On Error Resume Next
     If r > 0 And B > 0 Then LogN = Log(r) / Log(B)
End Function

'Add any function you want here:
'E.g.
Function ssqr(Number As Integer) As Double
'This is an integer accepting Sqr function. Uses For to calculate a number
On Error Resume Next
Dim n As Double, r As Double
Dim I As Byte
r = Number / 2
For I = 1 To 10
    r = (r + Number / r) / 2
Next I
ssqr = r
End Function

Function Factorial(ByVal Number As Integer) As Long
  On Error Resume Next
  'Gives the factorial of an integer
  Dim I%, Product&
  Product = 1
  If Number > 0 Then
    For I = 2 To Number Step 1
      Product = Product * I
    Next I
  Else
    For I = -1 To Number Step -1
      Product = Product * I
    Next I
  End If
  Factorial = Product
End Function

Function nCr(ByVal n%, ByVal r%) As Double
   'Returns the rth term of the nth line of pascals triangle
   On Error Resume Next
   nCr = Factorial(n) / (Factorial(r) * Factorial(n - r))
End Function
