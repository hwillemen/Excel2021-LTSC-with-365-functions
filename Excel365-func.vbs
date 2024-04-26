'VBA equivalent of structs:
Private Type ArrayInfo
    D1LB As Long    'Dimension 1: y or Rows
    D1UB As Long
    D1size As Long
    D2LB As Long    'Dimension 2: x or Columns
    D2UB As Long
    D2size As Long
    D3LB As Long    'Dimension 3: z or New x/y?
    D3UB As Long
    D3size As Long
End Type

Private Type ArraySize
    D1size As Long    'Dimension 1: y or Rows
    D2size As Long    'Dimension 2: x or Columns
    D3size As Long    'Dimension 3: z or New x/y?
End Type

Private Type DimInfo  'singe dimension + itterator
    LB As Long
    UB As Long
    size As Long
    N As Long
End Type

Private Type ArrayItter
    D1n As Long     'Dimension 1: y or Rows
    D2n As Long     'Dimension 2: x or Columns
    D3n As Long     'New X or Y Dimension
End Type

Public Const SIZE_RANGE As Long = -2
Public Const SIZE_NOT_ARRAY As Long = -1
Public Const SIZE_EMPTY As Long = 0

Public Const FORCE_1_DIM As Boolean = True

Public Sub OpenVBA()

Application.Goto Reference:="OpenVBA"

End Sub

'=========================================================================================
' Helper Functions: ArraySize, ToArray, TransposeArrayInfo, ResizeArray
'=========================================================================================

'Returns size:
'   -2 - A Range
'   -1 - Not an Array
'    0 - Empty
'  > 0 - Defined
Private Function ArraySize(ByVal arrValues As Variant, Optional ByVal dimensionOneBased As Long = 1) As Long
    Dim Result As Long: Result = SIZE_NOT_ARRAY 'Default to not an Array
    Dim LB As Long, UB As Long
  
    On Error GoTo NormalExit
  
    If (TypeOf arrValues Is Excel.Range) Then
        Result = SIZE_RANGE
    ElseIf (IsArray(arrValues) = True) Then
        Result = SIZE_EMPTY 'Move default to Empty
        LB = LBound(arrValues, dimensionOneBased) 'Possibly generates error
        UB = UBound(arrValues, dimensionOneBased) 'Possibly generates error
        If (LB <= UB) Then
            Result = UB - LB + 1 'Size greater than 1 or 1 (equals)
        End If
    End If
  
NormalExit:
    size = Result
End Function

'Returns array, even for single value
'size:
'   -1 - Not convertable to an Array
'    0 - Empty
'  > 0 - Size of array's 1st Dimension
'=ToArray(arrValues (1 value,array or Range), *size,*inf As ArrayInfo)
Private Function ToArray(ByVal arrValues As Variant, ByRef size As Long, _
    ByRef inf As ArrayInfo _
) As Variant
    'Note: in VBA it is not possible to make a custom type (struct) Optional input or cast it to an Optional Variant
    'extension idea if needed: add 2 optional Booleans; only_row1, only_col1
    Dim arrTemp As Variant, temp As Variant
    size = SIZE_NOT_ARRAY 'Default to not an Array
    
    On Error GoTo err
   
    inf.D1size = 1
    inf.D2size = 0
    If (TypeOf arrValues Is Excel.Range) Then
        If arrValues.Rows.Count * arrValues.Columns.Count = 1 Then
            ReDim arrTemp(1 To 1): arrTemp(1) = arrValues.Value
            size = 1: GoTo ExitF
        Else
            arrTemp = arrValues.Value
        End If
    ElseIf (IsArray(arrValues) = True) Then
        arrTemp = arrValues
    Else
        ReDim arrTemp(1 To 1): arrTemp(1) = arrValues
        size = 1: GoTo ExitF
    End If
    
    On Error GoTo ExitF
    size = SIZE_EMPTY 'Move default to Empty
    inf.D1LB = LBound(arrTemp, 1) 'Possibly generates error
    inf.D1UB = UBound(arrTemp, 1) 'Possibly generates error
    If (inf.D1LB <= inf.D1UB) Then
        size = inf.D1UB - inf.D1LB + 1 'size = 1 or greater
        inf.D1size = size
    End If
    inf.D2LB = LBound(arrTemp, 2) 'Maybe generates error
    inf.D2UB = UBound(arrTemp, 2) 'Maybe generates error
    If (inf.D2LB <= inf.D2UB) Then
        inf.D2size = inf.D2UB - inf.D2LB + 1 'size = 1 or greater
    End If
    inf.D3LB = LBound(arrTemp, 3) 'Maybe generates error
    inf.D3UB = UBound(arrTemp, 3) 'Maybe generates error
    If (inf.D3LB <= inf.D2UB) Then
        inf.D3size = inf.D3UB - inf.D3LB + 1 'size = 1 or greater
    End If

ExitF:
    ToArray = arrTemp
    Exit Function
err:
    size = SIZE_NOT_ARRAY
    MsgBox "ToArray() error " & err.Number & ": " & err.Description
    ToArray = Empty
End Function

Private Function TransposeArrayInfo(inf As ArrayInfo) As ArrayInfo
    Dim temp As Long
    temp = inf.D1LB: inf.D1LB = inf.D2LB: inf.D2LB = temp
    temp = inf.D1UB: inf.D1UB = inf.D2UB: inf.D2UB = temp
    temp = inf.D1size: inf.D1size = inf.D2size: inf.D2size = temp
    TransposeArrayInfo = inf
End Function

'Every dimension to 1-based array
Public Function ResizeArray(arrValues As Variant, ByVal NewSizeD1 As Long, _
    Optional ByVal NewSizeD2 As Long = 0, Optional ByVal NewSizeD3 As Long = 0 _
) As Variant
    Dim arrTemp As Variant, temp As Variant, size As Long
    Dim inf As ArrayInfo, i As ArrayItter, newi As ArrayItter
    
    'Resize new array and sanity check
    If NewSizeD3 > 0 Then
        If NewSizeD2 < 1 Or NewSizeD1 < 1 Then err.Description = "3D array with other dimension < 1": GoTo err
        ReDim arrTemp(1 To NewSizeD1, 1 To NewSizeD2, 1 To NewSizeD3)
    ElseIf NewSizeD2 > 0 Then
        If NewSizeD1 < 1 Then err.Description = "2D array with other dimension < 1": GoTo err
        ReDim arrTemp(1 To NewSizeD1, 1 To NewSizeD2)
    ElseIf NewSizeD1 > 0 Then
        ReDim arrTemp(1 To NewSizeD1)
    Else
        GoTo ExitF ' return empty
    End If
    
    'retrieve ArrayInfo and array if arrValues is a Range
    arrValues = ToArray(arrValues, size, inf)
    
    For i.D1n = inf.D1LB To (inf.D1LB + NewSizeD1 - 1)
        newi.D1n = i.D1n - inf.D1LB + 1
        If NewSizeD2 = 0 Then
            'new 1D:
            If i.D1n > inf.D1UB Then
                temp = Empty
            ElseIf inf.D3size > 0 Then
                temp = arrValues(i.D1n, inf.D2LB, inf.D3LB)
            ElseIf inf.D2size > 0 Then
                temp = arrValues(i.D1n, inf.D2LB)
            Else
                temp = arrValues(i.D1n)
            End If
            arrTemp(newi.D1n) = temp
        Else
            'new 2D:
            For i.D2n = inf.D2LB To (inf.D2LB + NewSizeD2 - 1)
                newi.D2n = i.D2n - inf.D2LB + 1
                
                If NewSizeD3 = 0 Then
                    If i.D2n > inf.D2UB Then
                        temp = Empty
                    ElseIf inf.D3size > 0 Then
                        temp = arrValues(i.D1n, i.D2n, inf.D3LB)
                    ElseIf inf.D2size > 0 Then
                        temp = arrValues(i.D1n, i.D2n)
                    Else
                        temp = arrValues(i.D1n)
                    End If
                    arrTemp(newi.D1n, newi.D2n) = temp
                Else
                    'new 3D:
                    For i.D3n = inf.D3LB To (inf.D3LB + NewSizeD3 - 1)
                        newi.D3n = i.D3n - inf.D3LB + 1
                            
                        If i.D3n > inf.D3UB Then
                            temp = Empty
                        ElseIf inf.D3size > 0 Then
                            temp = arrValues(i.D1n, i.D2n, i.D3n)
                        ElseIf inf.D2size > 0 Then
                            temp = arrValues(i.D1n, i.D2n)
                        Else
                            temp = arrValues(i.D1n)
                        End If
                        arrTemp(newi.D1n, newi.D2n, newi.D3n) = temp
                    Next
                End If
            Next
        End If
    Next

ExitF:
    ResizeArray = arrTemp
    Exit Function
err:
    MsgBox "ResizeArray() error " & err.Number & ": " & err.Description
    ResizeArray = Empty
End Function

Public Sub ResizeArrayTest()
    Dim arr() As Variant, temp As Variant
    Dim x As Long, y As Long
    
    'expand 1D to 2D
    ReDim arr(6 To 8)
    For x = 6 To 8
        arr(x) = CLng(Rnd() * 1000)
    Next
    temp = ResizeArray(arr, 3, 2)
    
    'convert 2D to 1D
    ReDim arr(6 To 8, 4 To 7)
    For x = 6 To 8
        For y = 4 To 7
            arr(x, y) = CLng(Rnd() * 1000)
        Next
    Next
    temp = ResizeArray(arr, 2)

    '2D, keep first column
    temp = ResizeArray(arr, 3, 1)
    
    'convert 2D to 3D
    temp = ResizeArray(arr, 3, 4, 2)
    temp = Empty
End Sub

'=========================================================================================
' Text Manipulation: TEXTSPLIT, TEXTBEFORE and TEXTAFTER
'=========================================================================================
'https://support.microsoft.com/en-us/Search/results?query=TEXTSPLIT+function
'https://support.microsoft.com/en-us/Search/results?query=TEXTBEFORE+function
'https://exceljet.net/functions/textbefore-function
'https://support.microsoft.com/en-us/Search/results?query=TEXTAFTER+function
'https://exceljet.net/functions/textafter-function

'=TEXTSPLIT(text,col_delimiter,[row_delimiter],[ignore_empty], [match_mode], [pad_with])
Public Function TEXTSPLIT(text As String, Optional col_delimiter As String = "", _
    Optional row_delimiter As String = "", Optional ignore_empty As Boolean = False, _
    Optional match_mode As Long = 0, Optional pad_with As Variant = CVErr(xlErrNA) _
) As Variant
'ignore_empty       Specify TRUE to ignore consecutive delimiters. Defaults to FALSE, which creates an empty cell.
'match_mode: Default 0 for case-sensitive match, 1 for case-insensitive.
    Dim arrOut As Variant, arrRows As Variant, arrTemp As Variant
    Dim inf As ArraySize, i As ArrayItter
    Dim temp As Variant
    'Map match_mode to a Split() CompareMethod, vbTextCompare is case-insensitive:
    match_mode = IIf(match_mode = 0, vbBinaryCompare, vbTextCompare)
    
    On Error GoTo err
    arrRows = Split(text, row_delimiter, , match_mode)
    If LBound(arrRows) <> 0 Then err.Description = "Split() did not return 0 based array! code needs adjusting.": GoTo err
    inf.D1size = UBound(arrRows) + 1
    If ignore_empty Then 'loop over rows and delete empty:
        err.Description = "ignore_empty not implemented yet!": GoTo err
    End If
    inf.D2size = 0
    
    ReDim arrTemp(0 To inf.D1size - 1)
    For i.D1n = 0 To inf.D1size - 1
        arrTemp(i.D1n) = Split(arrRows(i.D1n), col_delimiter, , match_mode)
        temp = UBound(arrTemp(i.D1n)) + 1
        If ignore_empty Then 'loop over cols and delete empty:
            err.Description = "ignore_empty not implemented yet!": GoTo err
        End If
        If temp > inf.D2size Then inf.D2size = temp
    Next
    
    'determine output type (string or 2D)
    If inf.D1size * inf.D2size = 1 Then
        TEXTSPLIT = text 'No delimiters found, return original
        Exit Function
    Else
        ReDim arrOut(inf.D1size - 1, inf.D2size - 1)
    End If
    
    'Copy values and fill empty:
    For i.D1n = 0 To inf.D1size - 1
        temp = UBound(arrTemp(i.D1n))
        For i.D2n = 0 To temp
            'If IsEmpty(arrTemp(i.D1n)(i.D2n)) Or arrTemp(i.D1n)(i.D2n) = "" Then
            If arrTemp(i.D1n)(i.D2n) = "" Then
                arrOut(i.D1n, i.D2n) = pad_with
            Else
                arrOut(i.D1n, i.D2n) = arrTemp(i.D1n)(i.D2n)
            End If
        Next
        For i.D2n = temp + 1 To inf.D2size - 1
            arrOut(i.D1n, i.D2n) = pad_with
        Next
    Next
    
ExitF:
    TEXTSPLIT = arrOut
    Exit Function
err:
    MsgBox "TEXTSPLIT() error " & err.Number & ": " & err.Description
    TEXTSPLIT = CVErr(xlErrValue)
End Function

'=TEXTBEFORE(text,delimiter,[instance_num], [match_mode], [match_end], [if_not_found])
Public Function TEXTBEFORE(text As String, Optional delimiter As String = "", _
    Optional ByVal instance_num As Long = 1, Optional ByVal match_mode As Long = 0, _
    Optional match_end As Long = 0, Optional if_not_found As Variant = CVErr(xlErrNA) _
) As Variant
    TEXTBEFORE = TEXTBEFORE_AFTER(True, text, delimiter, instance_num, match_mode, match_end, if_not_found)
End Function

'=TEXTAFTER(text,delimiter,[instance_num], [match_mode], [match_end], [if_not_found])
Public Function TEXTAFTER(text As String, Optional delimiter As String = "", _
    Optional ByVal instance_num As Long = 1, Optional ByVal match_mode As Long = 0, _
    Optional match_end As Long = 0, Optional if_not_found As Variant = CVErr(xlErrNA) _
) As Variant
    TEXTAFTER = TEXTBEFORE_AFTER(False, text, delimiter, instance_num, match_mode, match_end, if_not_found)
End Function

'=TEXTBEFORE/TEXTAFTER(text,delimiter,[instance_num], [match_mode], [match_end], [if_not_found])
Private Function TEXTBEFORE_AFTER(Before As Boolean, text As String, Optional delimiter As String = "", _
    Optional ByVal instance_num As Long = 1, Optional ByVal match_mode As Long = 0, _
    Optional match_end As Long = 0, Optional if_not_found As Variant = CVErr(xlErrNA) _
) As Variant
'instance_num: The instance of the delimiter before/after to extract. Negative from the end.
'match_mode: Default 0 for case-sensitive, 1 for case-insensitive.
'match_end  Treats the end of text as a delimiter. By default text is an exact match.
'if_not_found   Value returned on no match. #N/A by default
    Dim Result As Variant, DelimCount As Long, Positions() As Long, PosN As Long, PosPrev As Long
    
    On Error GoTo err
    
    'Long/Boolean to correct value mapping:
    match_mode = IIf(match_mode = 0, vbBinaryCompare, vbTextCompare)
    match_end = IIf(match_end = -1, 1, match_end)  'Excel TRUE translates to -1
    'Parameter check:
    If instance_num = 0 Or Abs(instance_num) > Len(text) Then
        Result = CVErr(xlErrValue): GoTo ExitF
    End If
    If Len(delimiter) = 0 Then
        If Before Then
            PosN = IIf(instance_num > 0, 1, Len(text) + 1): GoTo MatchFound
        Else
            PosN = IIf(instance_num > 0, 0, Len(text)): GoTo MatchFound
        End If
    End If

    DelimCount = (Len(text) - Len(Replace(text, delimiter, "", , , match_mode))) / Len(delimiter)
    'Convert instance_num to positive:
    If instance_num < 0 Then instance_num = DelimCount + instance_num + 1 '-> 2 + -1 = 1 so add 1
    
    'DelimCount/instance_num check and Positions() ReDim:
    If match_end = 0 Then
        If DelimCount = 0 Or instance_num < 1 Or instance_num > DelimCount Then
            Result = if_not_found: GoTo ExitF
        End If
        ReDim Positions(1 To DelimCount)
    Else 'match_end = 1
        If DelimCount = 0 Or instance_num < 0 Or instance_num > DelimCount + 1 Then
            'instance_num = -(DelimCount + 1) becomes 0 internally and is the start of the text
            Result = if_not_found: GoTo ExitF
        End If
        ReDim Positions(0 To DelimCount + 1) '0 and +1 positions are for match_end
        Positions(0) = IIf(Before, 1, 0)                                  'pos at start/before
        Positions(DelimCount + 1) = IIf(Before, Len(text) + 1, Len(text)) 'pos after/at end
    End If
    
    'Positions of delimiter
    PosPrev = 1
    For PosN = 1 To DelimCount 'InStr() cannot search from back so search full text
        Positions(PosN) = InStr(PosPrev, text, delimiter, match_mode)
        PosPrev = Positions(PosN) + 1
    Next
    
    PosN = Positions(instance_num)
MatchFound:
    If Before Then
        Result = Left(text, PosN - 1)
    Else
        Result = Right(text, Len(text) - PosN)
    End If
    
ExitF:
    TEXTBEFORE_AFTER = Result
    Exit Function
err:
    MsgBox "TEXT" & IIf(Before, "BEFORE", "AFTER") & "() error " & err.Number & ": " & err.Description
    TEXTBEFORE_AFTER = CVErr(xlErrValue)
End Function

'=========================================================================================
' Data Layout: TOCOL, TOROW, VSTACK, HSTACK
'=========================================================================================

'=TOCOL(array, [ignore], [scan_by_column])
Public Function TOCOL(arrRange As Variant, Optional ignore As Long = 0, _
    Optional scan_by_column As Boolean = False _
) As Variant
    TOCOL = TOCOL_ROW(True, arrRange, ignore, scan_by_column)
End Function

'=TOROW(array, [ignore], [scan_by_column])
Public Function TOROW(arrRange As Variant, Optional ignore As Long = 0, _
    Optional scan_by_column As Boolean = False _
) As Variant
    TOROW = TOCOL_ROW(False, arrRange, ignore, scan_by_column)
End Function

Private Function TOCOL_ROW(ToColumn As Boolean, arrRange As Variant, Optional ignore As Long = 0, _
    Optional scan_by_column As Boolean = False _
) As Variant
'ignore:0   Keep all values (default)
'       1   Ignore blanks
'       2   Ignore Errors
'       3   Ignore blanks And Errors
    Dim arrOut As Variant, arrTemp(), size_d1 As Long
    Dim inf As ArrayInfo, i As ArrayItter
    Dim temp As Variant
   
    On Error GoTo err
    
    arrTemp = ToArray(arrRange, size_d1, inf)
    If (inf.D1size * inf.D2size) = 0 Then arrOut = arrRange: GoTo ExitF
    
    'new array to maximum size:
    If ToColumn Then
        ReDim arrOut(1 To (inf.D1size * inf.D2size), 1 To 1)
    Else 'ToRow
        ReDim arrOut(1 To (inf.D1size * inf.D2size))
    End If
    'if specified, switch D1 and D2 while scanning
    If scan_by_column Then inf = TransposeArrayInfo(inf)
    
    i.D3n = 0
    For i.D1n = inf.D1LB To inf.D1UB
        For i.D2n = inf.D2LB To inf.D2UB
            If scan_by_column Then 'switch D1 and D2
                temp = arrTemp(i.D2n, i.D1n)
            Else
                temp = arrTemp(i.D1n, i.D2n)
            End If
            If (ignore = 1 Or ignore = 3) And IsEmpty(temp) Then GoTo NextItem
            If (ignore = 2 Or ignore = 3) And IsError(temp) Then GoTo NextItem
            If ToColumn Then
                arrOut(i.D3n + 1, 1) = temp
            Else 'ToRow
                arrOut(i.D3n + 1) = temp
            End If
            i.D3n = i.D3n + 1
NextItem:
        Next
    Next
    'Trim to correct length:
    If ToColumn Then
        arrOut = ResizeArray(arrOut, i.D3n, 1)
        'ReDim Preserve arrOut(1 To i.D3n, 1 To 1) 'only works on last dimension!
    Else 'ToRow
        ReDim Preserve arrOut(1 To i.D3n)
    End If
    
    
ExitF:
    TOCOL_ROW = arrOut
    Exit Function
err:
    MsgBox "TO" & IIf(ToColumn, "COL", "ROW") & "() error " & err.Number & ": " & err.Description
    TOCOL_ROW = CVErr(xlErrValue)
End Function

Public Function VSTACK(ParamArray Ranges() As Variant) As Variant
' Optional last argument "ForceOneDim = True": Transpose 2D Arrays/Ranges to 1D
    VSTACK = VHSTACK(True, Ranges())
End Function

Public Function HSTACK(ParamArray Ranges() As Variant) As Variant
' Optional last argument "ForceOneDim = True": Transpose 2D Arrays/Ranges to 1D
    HSTACK = VHSTACK(False, Ranges)
End Function

Private Function VHSTACK(Vertical As Boolean, ParamArray Ranges() As Variant) As Variant
' Optional last argument "ForceOneDim = True": Transpose 2D Arrays/Ranges to 1D
    Dim arrOut() As Variant, arrTemp As Variant, temp As Variant, size As Long
    Dim iRanges As DimInfo   'Ranges array dimension
    Dim inf() As ArrayInfo  'info of the ranges/arrays within Ranges
    Dim i As ArrayItter     'Iterators: D1=Rows, D2=Cols, D3=New Rows(VSTACK)/Cols(HSTACK)
    Dim infNew As ArraySize 'info of the new array: D1=RowCount, D2=ColCount
    Dim ForceOneDim As Boolean: ForceOneDim = False ' Optional Argument
    Dim Horizontal As Boolean: Horizontal = Not Vertical 'for code readability
    On Error GoTo err
    
    arrTemp = Ranges(0) 'unpack delegated ParamArray
    Ranges = arrTemp
    'Ranges info/dimension & ForceOneDim:
    iRanges.LB = LBound(Ranges): iRanges.UB = UBound(Ranges)
    If iRanges.LB <> 0 Then err.Description = "ParamArray Ranges() not 0 based! code needs checking.": GoTo err
    If (VarType(Ranges(iRanges.UB)) = vbBoolean) Then
        ForceOneDim = Ranges(iRanges.UB)
        iRanges.UB = iRanges.UB - 1
    End If
    iRanges.size = iRanges.UB + 1
    
    'Helpers:
    ReDim arrTemp(iRanges.LB To iRanges.UB)
    ReDim inf(iRanges.LB To iRanges.UB)
    
    'Detect new array dimensions, store arrays in arrTemp:
    For iRanges.N = iRanges.LB To iRanges.UB
        arrTemp(iRanges.N) = ToArray(Ranges(iRanges.N), size, inf(iRanges.N))
        If size = SIZE_NOT_ARRAY Then err.Description = "Range" & iRanges.N & " is neither a Range or an Array!"": GoTo err"

        'infNew D1: Rows, D2: Cols
        If Vertical And Not ForceOneDim Then
            infNew.D1size = infNew.D1size + inf(iRanges.N).D1size
            If infNew.D2size < inf(iRanges.N).D2size Then infNew.D2size = inf(iRanges.N).D2size
        ElseIf Horizontal And Not ForceOneDim Then
            If infNew.D1size < inf(iRanges.N).D1size Then infNew.D1size = inf(iRanges.N).D1size
            infNew.D2size = infNew.D2size + inf(iRanges.N).D2size
        ElseIf Vertical And ForceOneDim Then    ' Transposing 2D arrays:
            infNew.D1size = infNew.D1size + (inf(iRanges.N).D1size * inf(iRanges.N).D2size)
            infNew.D2size = 1
        ElseIf Horizontal And ForceOneDim Then  ' Transposing 2D arrays:
            infNew.D1size = infNew.D1size + (inf(iRanges.N).D1size * inf(iRanges.N).D2size)
            infNew.D2size = 0
        End If
    Next 'iRanges.N

    'Set new array dimensions:
    If infNew.D2size = 0 Then                 ' input were 1D arrays or 2D with "ForceOneDim"
        ReDim arrOut(1 To infNew.D1size)
    Else
        ReDim arrOut(1 To infNew.D1size, 1 To infNew.D2size)
    End If
    i.D3n = 1

    If Horizontal Then GoTo H_STACK
V_STACK:
    For iRanges.N = iRanges.LB To iRanges.UB
        For i.D1n = 1 To inf(iRanges.N).D1size      'Rows
            For i.D2n = 1 To inf(iRanges.N).D2size  'Columns
                If ForceOneDim Then 'Transposing to 1 column
                    arrOut(i.D3n, 1) = arrTemp(iRanges.N)(i.D1n, i.D2n)
                    i.D3n = i.D3n + 1
                Else
                    arrOut(i.D3n, i.D2n) = arrTemp(iRanges.N)(i.D1n, i.D2n)
                End If
            Next
            If Not ForceOneDim Then i.D3n = i.D3n + 1 ' Regular 1D or 2D array
        Next
    Next 'iRanges.N
    GoTo ExitF
    
H_STACK:
    For iRanges.N = iRanges.LB To iRanges.UB
        'HSTACK: scan_by_column to not overflow i.D3n
        For i.D2n = 1 To inf(iRanges.N).D2size      'Columns
            For i.D1n = 1 To inf(iRanges.N).D1size  'Rows
                temp = arrTemp(iRanges.N)(i.D1n, i.D2n)
                If ForceOneDim Then 'Transposing to 1 row
                    arrOut(i.D3n) = temp
                    i.D3n = i.D3n + 1
                Else
                    arrOut(i.D1n, i.D3n) = temp
                End If
            Next
            If Not ForceOneDim Then i.D3n = i.D3n + 1 ' Regular 1D or 2D arrays
        Next
     Next 'iRanges.N
    
ExitF:
    VHSTACK = arrOut
    Exit Function
err:
    MsgBox IIf(Vertical, "V", "H") & "STACK() error " & err.Number & ": " & err.Description
    VHSTACK = CVErr(xlErrValue)
End Function

Public Sub TestVSTACK()
    Dim MyRange1 As Variant, MyRange2 As Variant, Result1 As Variant, Result2 As Variant
    'Set MyRange1 = Range("$AB$20:$AB$50")
    Set MyRange1 = Range("$AB$20:$AC$50")
    MyRange2 = Array()
    'Set MyRange2 = Range("$AC$20:$AC$50")
    'MyRange1 = MyRange1.Value
    'MyRange2 = MyRange2.Value
    
    If ArraySize(MyRange2) = SIZE_EMPTY Then
        Result1 = VSTACK_Clone(MyRange1, FORCE_1_DIM)
        Result2 = SortUnique(MyRange1)
    Else
        Result1 = VSTACK_Clone(MyRange1, MyRange2, FORCE_1_DIM)
        Result2 = SortUnique(Result1)
    End If
    MsgBox "Done"
    
End Sub

'=========================================================================================
' Organizing Data: WRAPROWS, WRAPCOLS
'=========================================================================================

'=========================================================================================
' Data Management: TAKE, DROP, CHOOSECOLS, CHOOSEROWS, EXPAND
'=========================================================================================

'=========================================================================================
' Other Data Functions: SortUnique
'=========================================================================================

Public Function SortUnique(Range As Variant, Optional exactly_once As Boolean = False) As Variant
' Takes care of filtering blanks, unique and sorting of a Range/Array.
' Equivalent of =SORT(UNIQUE(FILTER(Range; Range<>""))) but with only 1x Range input.
' Better for readability and performance when Range is the output of another function.
End Function