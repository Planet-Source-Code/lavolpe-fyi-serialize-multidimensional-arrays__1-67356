VERSION 5.00
Begin VB.Form frmTestHacks 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   375
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Results are written to the debug window. Press Ctrl + G to view that window."
      Height          =   660
      Left            =   615
      TabIndex        =   1
      Top             =   1605
      Width           =   2940
   End
End
Attribute VB_Name = "frmTestHacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
' ^^ API limitation. Cannot be used to return pointer to a string array. See http://support.microsoft.com/kb/199824

' See cArraySerialization for full description and discussion
' For example on using above VarPtrArray API, see the TestDates & TestObjects examples

Private cSerialization As New cArraySerialization

Private Sub Command1_Click()

    ' Purpose: Test different array data types
    
    Dim Looper As Long
    
    For Looper = 1 To 12
        Select Case Looper
        Case 1: Call TestBooleans
        Case 2: Call TestBytes
        Case 3: Call TestCurrency
        Case 4: Call TestDates
        Case 5: Call TestDoubles
        Case 6: Call TestIntegers
        Case 7: Call TestLongs
        Case 8: Call TestSingles
        Case 9: Call TestStrings
        Case 10: Call TestStaticStrings
        Case 11: Call TestObjects   ' Limited support. See cSerialization.SerializeArray
        Case 12: Call TestVariants  ' Limited support. See cSerialization.SerializeArray
        End Select
    Next
End Sub


Private Sub TestStrings()

    ' Per MSDN :: cannot use VarPtrArray on string arrays.
    ' http://support.microsoft.com/kb/199824

    Dim a() As String
    Dim p() As Byte
    Dim X As Long, Y As Long
    Dim I As Integer
    
    Debug.Print: Debug.Print "Testing Variable Length Strings..."
    Debug.Print " before:  ";
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            I = Int(Rnd * 20) + 1
            If I Mod 5 = 0 Then
                Debug.Print "[vbNullString] ";
            Else
                a(X, Y) = String(I, " ")
                For I = 1 To I
                    Mid$(a(X, Y), I, 1) = Chr$(Int(Rnd * 26) + 65)
                Next
                Debug.Print a(X, Y); " ";
            End If
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            If a(X, Y) = vbNullString Then
                Debug.Print "[vbNullString] ";
            Else
                Debug.Print a(X, Y); " ";
            End If
        Next
    Next
    Debug.Print
    
End Sub

Private Sub TestBytes()

    Dim a() As Byte
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Bytes..."
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            a(X, Y) = Int(Rnd * 255)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print

End Sub

Private Sub TestIntegers()

    Dim a() As Integer
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Integers..."
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            a(X, Y) = Int(Rnd * 32000)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print

End Sub

Private Sub TestLongs()

    Dim a() As Long
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Longs..."
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            a(X, Y) = Int(Rnd * vbWhite)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print

End Sub

Private Sub TestSingles()

    Dim a() As Single
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Singles..."
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            a(X, Y) = Int(Rnd * vbWhite)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print

End Sub

Private Sub TestBooleans()

    Dim a() As Boolean
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Booleans..."
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            If (Int(Rnd * vbWhite) Mod 2) = 0 Then a(X, Y) = True
            Debug.Print a(X, Y); " ";
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Erase a
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print a(X, Y); " ";
        Next
    Next
    Debug.Print

End Sub

Private Sub TestDoubles()

    Dim a() As Double
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Doubles..."
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            a(X, Y) = Int(Rnd * vbWhite)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print

End Sub

Private Sub TestCurrency()

    Dim a() As Currency
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Currency..."
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            a(X, Y) = Int(Rnd * vbWhite)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print a(X, Y);
        Next
    Next
    Debug.Print

End Sub

Private Sub TestDates()

    Dim a() As Date
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Dates..."
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            a(X, Y) = DateAdd("n", CLng(Rnd * vbWhite), Date)
            Debug.Print Format(a(X, Y), "ddmmmyyyy-HH:nn"); " ";
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), , VarPtrArray(a()), VarType(a)
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print Format(a(X, Y), "ddmmmyyyy-HH:nn"); " ";
        Next
    Next
    Debug.Print

End Sub

Private Sub TestObjects()

    Dim a() As Object
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Objects... All should be Nothing"
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            Debug.Print "[Nothing] ";
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), , VarPtrArray(a), VarType(a)
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            If a(X, Y) Is Nothing Then
                Debug.Print "[Nothing] ";
            Else
                Debug.Print "error objects are not encoded correctly"
            End If
        Next
    Next
    Debug.Print

End Sub

Private Sub TestVariants()
    
    Dim a() As Variant
    Dim p() As Byte
    Dim X As Long, Y As Long
    
    Debug.Print: Debug.Print "Testing Variants... All should be Empty"
    Debug.Print " before:  ";
    
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            If a(X, Y) = Empty Then Debug.Print "[Empty] ";
        Next
    Next
    Debug.Print
    
    cSerialization.SerializeArray p(), a()
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    Erase a
    a = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            If a(X, Y) = Empty Then
                Debug.Print "[Empty] ";
            Else
                Debug.Print "error variant arrays are not encoded correctly"
            End If
        Next
    Next
    Debug.Print

End Sub

Private Sub TestStaticStrings()

    ' FYI:
    ' Arrays containing variable length strings are actually pointers to memory where the strings can be found
    ' Arrays containing static strings are the actual strings

    Dim a() As String * 10  ' << static, fixed length
    Dim ss() As String
    Dim p() As Byte
    Dim X As Long, Y As Long
    Dim I As Integer
    
    Debug.Print: Debug.Print "Testing Static Strings (10 characters each)..."
    Debug.Print " before:  ";
    ReDim a(0 To 2, 5 To 8)
    For X = LBound(a, 1) To UBound(a, 1)
        For Y = LBound(a, 2) To UBound(a, 2)
            For I = 1 To 10
                Mid$(a(X, Y), I, 1) = Chr$(Int(Rnd * 26) + 65)
            Next
            Debug.Print a(X, Y); " ";
        Next
    Next
    Debug.Print
    
    ' notice the different parameter usage below...
    ' 1) Fixed length can't be assigned to Variants, therefore, can't pass the a() array in the Variant parameter
    ' 2) Using VarType on just the array causes errors (i.e., VarType(a()) is an error)
    ' 3) Because of that ^^, we can't get what the function needs: vbArray Or vbString. So we just force it
    cSerialization.SerializeArray p(), , VarPtrArray(a), vbString Or vbArray
    Debug.Print vbTab; "multidimensional array converted to single byte array of "; UBound(p) + 1; " bytes"
    
    Erase a
    
    ' also when returning the array, we can't assign the result to a fixed length string array else errors.
    ' To get the strings into the fixed array, looping the returned dynamic array and assigning each dynamic array string
    ' to the static array string would be required.
    
    ' Bottom line :: fixed length arrays are not truly compatible. They can be serialized, but can't be restored directly
    ss = cSerialization.DeSerializeArray(p())
    
    Debug.Print "restored: ";
    For X = LBound(ss, 1) To UBound(ss, 1)
        For Y = LBound(ss, 2) To UBound(ss, 2)
            Debug.Print ss(X, Y); " ";
        Next
    Next
    Debug.Print
    
End Sub

Private Sub Form_Load()
    Randomize Timer
End Sub


