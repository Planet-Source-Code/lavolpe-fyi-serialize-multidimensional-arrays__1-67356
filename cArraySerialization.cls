VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArraySerialization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' http://www.codeguru.com/vb/gen/vb_misc/algorithms/article.php/c7495
' ^^ interesting read

' The results of some experimenting....

' The problem I faced was storing an array of Longs or Strings into the property bag of a user control.
' Anyone that is familiar with user controls, storage of arrays is poorly supported

' Well I didn't want to store each array element as a separate property. I figured there had to be another way.
' I could convert any array into a byte array and store it, but how do I get a dimensional array into a byte array,
'    then how do I get it from a byte array back to the dimensional array keeping the same array data type and also
'    keeping the same LBounds and UBounds?

' The 1st issue: getting a multidimensional array into a byte array.
'   The answer was pretty easy, calculate the number of elements of an array then multiply by the number of bytes
'   a single element uses. After that, create a byte array of correct size and simply use CopyMemory. This was
'   rather easy but String arrays not that simple, but still could be done converting strings to byte arrays (StrConv).

' The 2nd issue: converting a byte array back to another data type array or string array.
'   The initial answer was to create uninitialized arrays for all possible data types, using the data type of the
'   array before the array was converted to bytes, then set a variant to that unitialized array. Kind of like:
'       Select Case OriginalDataType
'       Case vbLong:  Dim myVariant() As Long
'       ... etc
'   Ok, how do we redim myVariant if the original array was multidimensional?  Hmmmmm.
'       The initial answer was to cache the LBounds & UBounds of the original array, then using another Select Case tree,
'       dimension the array appropriately.  Kind of like:
'           Select Case nrDimensions
'           Case 1: ReDim MyVariant(LB1 to UB1)
'           Case 2: ReDim MyVariant(LB1 to UB1, LB2 to UB2)
'           Case 3: ReDim MyVariant(LB1 to UB1, LB2 to UB2, LB3 to UB3)
'           ... etc

' The above logic was my first attempt, and I became frustrated knowing there had to be a way to do this without having
' to restrict dimensions to how many Select Case statements I felt was appropriate, or without having to use a lengthy
' Select Case statement to establish the correct data type.

' Thank goodness for patience, a little knowledge on memory pointers, and the URL noted at top of this page

' Enclosed are the results.
' It works a little like this...

' Converting multidimensional array to bytes
' 1) Using the original array, cache its LBounds & UBounds, its variable type, and number of dimensions
' 2) Also cache the contents of the array. Object arrays can't be cached, objects need to be serialized independently
'     -- Variable length strings require a bit more work because we need the length of each string too

' Converting back to multidimensional array
' 1) Use a variant variable
' 2) Hack it to make VB think it is of the variable type we want (i.e., Long, String, Object, etc)
' 3) Create a SafeArray that points into our cached data array (our byte array)
' 4) Hack the variant again to make VB think that it is arrayed and its SafeArray pointer is the one we created
' 5) Assign the needed array variable or another Variant to our hacked variant.
'       -- In the assignment, VB will copy what it thinks is the contents of the hacked variant (really our cached array data)
'           into the newly assigned variable. This allows us to free our cached data which will go out of scope anyway.
' 6) Release our cached data, unhack the variant, destroy any temporary Pointer arrays that may have been created
'    -- String & Object arrays are a bit more complicated because the array is really an array of pointers to pointers

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Function SerializeArray(outArray() As Byte, Optional inArray As Variant, _
                    Optional ByVal ArrayPointer As Long, Optional ByVal ArrayVariableType As VariantTypeConstants) As Boolean

    ' PARAMETERS............................
    ' outArray() :: the returned serialized byte array.
    ' inArray :: your array, the one you want serialized into outArray
    '   -- if not provided, then ArrayPointer parameter is used
    ' ArrayPointer :: the array's pointer, use VarPtrArray() API.
    '   -- if not provided then inArray parameter is used
    ' ArrayVariableType: if ArrayPointer is used, this must also be provided
    '   -- use VarType(myArray)
    ' Return Value is True if no errors are encountered.
    
    ' Which to use?. Except for string arrays, Preferrably you want to use VarPtrArray() and VarType parameters
    ' Why? Passing your array via the variant inArray parameter forces VB to copy
    ' your array. We will be copying your array too -- 3 copies!
    ' But by passing the ArrayPointer & ArrayVariableType parameters, no extra copying.
    ' -- For string arrays, VarPtrArray API won't work. See http://support.microsoft.com/kb/199824
    
    ' Ex: Call SerializeArray(arrayBytes(), myArray()) << passing your array to variant parameter
    ' Ex: Call SerializeArray(arrayBytes(), , VarPtrArray(myArray), VarType(myArray)) << passing without using variant parameter

    '   Note: a max of 2,147,483,647 bytes can be used to serialize the array. But in reality
    '       any byte array nearly that large would probably generate out of memory errors.
    '   Note: This routine could be modified to provide a starting array element and
    '         number of elements to serialize if an array is expected to be extremely huge.
    '         Then this routine could be called in steps to serialize sections of the array.
    '   Note: Byte arrays, regardless of number of dimensions, do not have to be serialized to be added to a property bag object.
    '         The property bag stores the byte arrays along with its SafeArray descriptor.
    
    Dim saFarPtr As Long        ' pointer to memory, reusable
    Dim saPtr As Long           ' pointer to passd array's SafeArray structure
    Dim stringPtr As Long       ' pointer to strings when processing string arrays
    Dim nSize As Long           ' size of a data element or string
    Dim nrDims As Integer       ' number of dimensions in the array
    Dim nrElements As Long      ' total elements in all dimensions
    Dim lValue As Long          ' resuable value variable

    ' the outArray is formatted as follows:
    '   bytes 0 & 1: nr of dimensions with &H4000 added
    '   bytes 2 & 3: variable type with &H4000 added -- will included the vbArray byte also
    '   bytes 4 to 4+8*dimensions: the number of elements for each dimensions & each LBound
    '   remaining bytes as described below.
    ' -- For numbers, booleans, & dates
    '       remaining bytes: the actual array data
    ' -- For Strings
    '       remaining bytes: each string as:
    '       : 4 bytes for length of string
    '       : If string is not vbNullString, then the 2-byte-per-character string
    ' -- For Objects and Variants
    '       no additional data. Objects are not stored; only the array dimensions are retained
    '       You should serialize objects using whatever methods you are familiar with
    ' -- For Variant arrays
    '       no additional data. only the array dimensions are retained.
    '       You should walk the array to serialize each item in the array
    '       For X = LBound(vArray) to UBound(vArray)
    '           If (VarType(vArray(x)) = (vbArray or vbVariant)) Then
    '               Call SerializeArray(serializedBytes(), vArray(x))
    '           End If
    '       Next
    
    On Error GoTo ExitRoutine
    
    ' get some core info for/from the safe array contained inside the variant
    If Not ArrayPointer = 0 Then                                    ' user passed VarPtrArray()
        If (ArrayVariableType And vbArray) = 0 Then Exit Function   ' ensure the vbArray flag is included
        saFarPtr = ArrayPointer
    Else
        ArrayVariableType = VarType(inArray)
        If (ArrayVariableType And vbArray) = 0 Then Exit Function   ' ensure the vbArray flag is included
        CopyMemory saFarPtr, SafeOffset(VarPtr(inArray), 8), 4&     ' pointer to SafeArray pointer (far pointer)
        CopyMemory lValue, inArray, 4&                  ' since passed by Ref, we need to adjust the far pointer
        If (lValue And &H4000) = &H4000 Then CopyMemory saFarPtr, ByVal saFarPtr, 4&
   End If
    
    CopyMemory saPtr, ByVal saFarPtr, 4&            ' get the SafeArray pointer
    ' FYI: if a SafeArray structure does exist, then 4 bytes before the SafeArray (saPtr-4) is the VarType of the array
    
    If saPtr = 0 Then
        ReDim outArray(0 To 3)                      ' uninitialized array
        
    Else
        CopyMemory nrDims, ByVal saPtr, 2&               ' nr of dimensions
        CopyMemory nSize, ByVal SafeOffset(saPtr, 4), 4& ' size of data elements (1=byte,2=integer/boolean,4=long/single,8=double/date/currency)
                                                         ' for fixed length string arrays, nSize=len(String). For Variants, nSize=16
        ' calculate total number of elements contained in the array, include all dimensions
        nrElements = 1
        For nrDims = 0 To nrDims - 1    ' first dimension is found 16 bytes from header
            CopyMemory lValue, ByVal SafeOffset(saPtr, 16 + (8 * nrDims)), 4&
            nrElements = nrElements * lValue
        Next
        
        Select Case ArrayVariableType And Not vbArray
        
        Case vbString
        
            CopyMemory saFarPtr, ByVal SafeOffset(saPtr, 12), 4&    ' get pointer to 1st data element
            CopyMemory lValue, ByVal SafeOffset(saPtr, 2), 2&       ' get features; testing for fixed string vs variable length
            
            If Not (lValue And &H100) = &H100 Then                  ' it is fixed length of nSize
                ReDim outArray(0 To (nrDims * 8) + ((nSize + 4) * nrElements) + 3) ' size array
                lValue = nrDims * 8 + 4                             ' set position for placement in our array
                ' Since we know the length of every string, this is somewhat simpler, but arrays of fixed length strings are rarely used
                For nrElements = 0 To nrElements - 1
                    CopyMemory outArray(lValue), nSize, 4&                  ' cache size of array
                    CopyMemory outArray(lValue + 4), ByVal saFarPtr, nSize  ' cache string data
                    lValue = lValue + nSize + 4             ' position for next string
                    saFarPtr = SafeOffset(saFarPtr, nSize)  ' postion for next string pointer
                Next
                
            Else            ' variable length strings
                ' 2 passes: Pass 1: calculate size of array needed
                nSize = 0
                For nrElements = 0 To nrElements - 1
                    CopyMemory stringPtr, ByVal saFarPtr, 4&    ' get address of next array element (StrPtr)
                    If Not stringPtr = 0 Then                       ' test for vbNullString
                        CopyMemory lValue, ByVal SafeOffset(stringPtr, -4), 4&   ' get nr of bytes for string
                        nSize = nSize + lValue                      ' tally nr of bytes
                    End If
                    saFarPtr = SafeOffset(saFarPtr, 4)              ' adjust memory pointer position
                Next
                ' Pass 2: populate our array
                ReDim outArray(0 To (nrDims * 8) + nSize + (nrElements * 4&) + 3) ' resize array
                lValue = nrDims * 8 + 4                             ' set position for placement in our array
                CopyMemory saFarPtr, ByVal SafeOffset(saPtr, 12), 4& ' get pointer to the string array pointers
                For nrElements = 0 To nrElements - 1
                    CopyMemory stringPtr, ByVal saFarPtr, 4&        ' get address of array element (StrPtr)
                    If stringPtr = 0 Then                           ' test for vbNullString
                        lValue = lValue + 4                         ' position for next string
                    Else
                        stringPtr = SafeOffset(stringPtr, -4)
                        CopyMemory nSize, ByVal stringPtr, 4&       ' get nr of bytes for string
                        CopyMemory outArray(lValue), ByVal stringPtr, nSize + 4& ' copy the length of string & string data
                        lValue = lValue + nSize + 4                 ' position for next string
                    End If
                    saFarPtr = SafeOffset(saFarPtr, 4)  ' postion for next string pointer
                Next
            End If
            
        Case vbObject, vbVariant, vbDataObject, vbError, vbUserDefinedType
            ReDim outArray(0 To (nrDims * 8) + 3)               ' the data cannot be saved if it contains objects, variants should be walked as noted at top
                                                                ' data must be serialized or saved in some other fashion
                                                                ' only the array dimensions are cached for resizing array when restored
        Case Else ' numbers, dates, booleans
            ReDim outArray(0 To (nrDims * 8) + nrElements * nSize + 3)  ' resize array
            CopyMemory saFarPtr, ByVal SafeOffset(saPtr, 12), 4&        ' get pointer to 1st data element
            CopyMemory outArray(nrDims * 8 + 4), ByVal saFarPtr, nrElements * nSize ' copy array values
        End Select
        CopyMemory outArray(4), ByVal SafeOffset(saPtr, 16), nrDims * 8 ' cache nr elements and LBound of each dimension

    End If

    ' finish saving values into the byte array. Add &H4000 to for a valiation check when deserializing
    CopyMemory outArray(0), nrDims Or &H4000, 2&            ' cache dimensions
    CopyMemory outArray(2), ArrayVariableType Or &H4000, 2& ' cache VarType
    
    SerializeArray = True                     ' success
    
ExitRoutine:
End Function

Public Function DeSerializeArray(byteArray() As Byte) As Variant

    ' PURPOSE: Deserialize a previously serialized array.
    
    ' How this works is pretty simple, but very complicated at the same time....
    
    ' The bottom line is that we will abuse a Variant variable, telling the variant it is of a
    '   variable type we choose and that it is arrayed in the dimensions and bounds that we choose.
    '   The sneaky thing is that the variant is never dimensioned nor declared a specific data type.
    
    '   By setting this Function's return value (a Variant) to our abused Variant, we force
    '   VB to copy the array data from our local memory, via misdirection, so it will be preserved
    '   in the return variable.
    
    '   In fact, we could be dealing with arrays in all sorts of dimensions, having any number of
    '   LBound & UBounds and having every conceivable data type. But you won't see us trying to
    '   dimension the return variant or trying to calculate its individual dimensions LBounds/UBounds.
    
    ' For arrays of numbers, dates & booleans
    '   - The byteArray() contains everything we need: bounds for each dimension and the actual data
    '   - A soft SafeArray is created that points into the byteArray(). The abused variant points to that SafeArray
    ' For arrays of Objects
    '   - The byteArray() contains only dimensions & LBound/UBound information.
    '   - A soft SafeArray is created that points into another Long Array of empty pointers. The abused variant points to that SafeArray
    '   - The Object Array will be dimensioned, but every object is Nothing
    ' For arrays of Strings
    '   - The byteArray contains array dimensions & the strings along with the length of each string
    '   - A soft SafeArray is created that points into another Long Array of pointers. The abused variant points to that SafeArray
    '   - Each pointer/item in that Long Array will point to a position within the byteArray
    ' For arrays of Variants
    '   - The byteArray() contains only dimensions & LBound/UBound information.
    '   - A soft SafeArray is created that points into another Long Array of empty pointers. The abused variant points to that SafeArray
    '   - Each of the items in the Variant array should have been serialized individually by you.
    '   - See SerializeArray

    Dim saFarPtr As Long        ' pointer to memory address, reusable in this routine
    Dim nSize As Long           ' size of data element or size of any array
    Dim nrDims As Integer       ' number of dimensions in the array
    Dim nrElements As Long      ' number of elements in the entire array
    Dim lValue As Long          ' temp variable, reusable
    Dim tVariant As Variant     ' our return variant host
    Dim vType As Integer        ' variable type of the passed array
    Dim tSafeArray() As Byte    ' a soft SafeArray structure
    Dim tPointers() As Long     ' pointers, used when serializing string, variant & object arrays
    
    On Error GoTo ExitRoutine
    If UBound(byteArray) < 3 Or LBound(byteArray) < 0 Then Exit Function

    CopyMemory nrDims, byteArray(0), 2&             ' get number of dimensions from cache
    CopyMemory vType, byteArray(2), 2&              ' get array data type
    
    ' was created by SerializeArray?
    If (vType And (vbArray Or &H4000)) = 0 Then
        Exit Function    ' not an array created by SerializeArray
    ElseIf (nrDims And &H4000) = 0 Then
        Exit Function    ' not an array created by SerializeArray
    End If
    nrDims = nrDims And Not &H4000
    vType = vType And Not &H4000
        
    
    If Not nrDims = 0 Then  ' elseif nrDims=0, then we have an uninitialized array
        ReDim tSafeArray(1 To 16 + (nrDims * 8))    ' build a SafeArray descriptor
    
        CopyMemory tSafeArray(1), nrDims, 2&        ' SafeArray 1:2 > number of dimensions
                                                    ' SafeArray 3:4 > VT flag, not used for what we are doing
                                                    
        Select Case vType And Not vbArray
        Case vbString
            ' String arrays are quite different. The data within the array are actually pointers to other memory addresses
            ' When the pointer is zero, we have a vbNullString.
            ' Otherwise, the pointer points to an address where the string data begins
            ' & 4 bytes immediately preceeding that address contains the number of unicode bytes in the string
            
            CopyMemory tSafeArray(5), 4&, 4&        ' SafeArray 5:8 > number of bytes per array element (4 for pointers)
                                                    ' SafeArray 9:12 > locks, not used for what we are doing
            
            ' calculate the number of elements stored with the array
            nrElements = 1
            For nrDims = 0 To nrDims - 1
                CopyMemory lValue, byteArray(4 + nrDims * 8), 4&
                nrElements = nrElements * lValue
            Next
            ReDim tPointers(0 To nrElements - 1)    ' resize pointer array
            
            lValue = nrDims * 8 + 4                     ' position in array for 1st string
            For nrElements = 0 To nrElements - 1
                CopyMemory nSize, byteArray(lValue), 4& ' get size of string
                If Not nSize = 0 Then                   ' if not vbNullString then set pointer to location in the array for that string
                    CopyMemory tPointers(nrElements), VarPtr(byteArray(lValue + 4)), 4&
                End If
                lValue = lValue + nSize + 4             ' increment for next string position
            Next
            CopyMemory tSafeArray(13), VarPtr(tPointers(0)), 4& ' SafeArray 13:16 > The pointer to the array of data (or to other pointers for strings/objects/variants)
            
        Case vbObject, vbVariant, vbDataObject, vbError, vbUserDefinedType
            ' calculate the number of elements stored with the array
            nrElements = 1
            For nrDims = 0 To nrDims - 1
                CopyMemory lValue, byteArray(4 + nrDims * 8), 4&
                nrElements = nrElements * lValue
            Next
            ReDim tPointers(0 To nrElements - 1)    ' resize pointer array
            CopyMemory tSafeArray(5), 4&, 4&        ' SafeArray 5:8 > number of bytes per array element (4 for pointers)
                                                    ' SafeArray 9:12 > locks, not used for what we are doing
            CopyMemory tSafeArray(13), VarPtr(tPointers(0)), 4& ' SafeArray 13:16 > The pointer to the array of data (or to other pointers for strings/objects/variants)
            
        Case Else
            Select Case (vType And Not vbArray)     ' SafeArray 5:8 > number of bytes per array element
                Case vbLong, vbSingle               ' SafeArray 9:12 > locks, not used for what we are doing
                    CopyMemory tSafeArray(5), 4&, 4&
                Case vbCurrency, vbDate, vbDouble
                    CopyMemory tSafeArray(5), 8&, 4&
                Case vbInteger, vbBoolean
                    CopyMemory tSafeArray(5), 2&, 4&
                Case vbByte
                    CopyMemory tSafeArray(5), 1&, 4&
            End Select
                                                  
            CopyMemory tSafeArray(13), VarPtr(byteArray(nrDims * 8 + 4)), 4& ' SafeArray 13:16 > The pointer to the array of data (or to other pointers for strings/objects/variants)
            
        End Select
        
        CopyMemory tSafeArray(17), byteArray(4), nrDims * 8 ' SafeArray 17: 8*dimensions > the number of elements in dimension and the LBound of the dimension
        CopyMemory ByVal SafeOffset(VarPtr(tVariant), 8), VarPtr(tSafeArray(1)), 4&  ' Hack 1: Make tVariant appear to be arrayed with our data. 4bytes: pointer to our safearray
    
    End If
    
    ' using an empty Variant to host our makeshift array and safearray structure
    CopyMemory ByVal VarPtr(tVariant), vType, 2&    ' Hack 2: Tell tVariant what data type it is. Hack 1 is just a few lines up
    
    DeSerializeArray = tVariant
    CopyMemory ByVal VarPtr(tVariant), 0, 2&   ' undo what we did to the host Variant
    CopyMemory ByVal SafeOffset(VarPtr(tVariant), 8), 0&, 4&
    
    'Erase tSafeArray() ' erased when goes out of scope
    'Erase tPointers()  ' erased when goes out of scope
    
ExitRoutine:
End Function


Private Function SafeOffset(ByVal Ptr As Long, Offset As Long) As Long

    ' ref http://support.microsoft.com/kb/q189323/ ' unsigned math
    ' Purpose: Provide a valid pointer offset
    
    ' If a pointer +/- the offset wraps around the high bit of a long, the
    ' pointer needs to change from positive to negative or vice versa.
    
    ' A return of zero indicates the offset exceeds the min/max unsigned long bounds
    
    Const MAXINT_4NEG As Long = -2147483648#
    Const MAXINT_4 As Long = 2147483647
    
    If Offset = 0 Then
        SafeOffset = Ptr
    Else
    
        If Offset < 0 Then ' subtracting from pointer
            If Ptr < MAXINT_4NEG - Offset Then
                ' wraps around high bit (backwards) & changes to Positive from Negative
                SafeOffset = MAXINT_4 - ((MAXINT_4NEG - Ptr) - Offset - 1)
            ElseIf Ptr > 0 Then ' verify pointer does not wrap around 0 bit
                If Ptr > -Offset Then SafeOffset = Ptr + Offset
            Else
                SafeOffset = Ptr + Offset
            End If
        Else    ' Adding to pointer
            If Ptr > MAXINT_4 - Offset Then
                ' wraps around high bit (forward) & changes to Negative from Positive
                SafeOffset = MAXINT_4NEG + (Offset - (MAXINT_4 - Ptr) - 1)
            ElseIf Ptr < 0 Then ' verify pointer does not wrap around 0 bit
                If Ptr < -Offset Then SafeOffset = Ptr + Offset
            Else
                SafeOffset = Ptr + Offset
            End If
        End If
    End If

End Function


