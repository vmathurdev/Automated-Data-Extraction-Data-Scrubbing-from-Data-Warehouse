Sub runsql()

Dim sqltbl As Range, i As Long, usews As Worksheet, r As Long, c As Long
Dim strconnect As String, strquery As String, targetfield As String, targetfield2 As String, targetsht As String, strlabel2 As String
Dim arr As Variant, userng As Range, userng2 As Range
Dim rsdata As ADODB.Recordset
Dim cn As ADODB.Connection
Dim rscon As Object
'Dim accApp As Access.Application
Dim shtnm As String, wbnm As String, addr As String
Dim conn As Connection

Dim strtyp As String
i = 1



Set sqltbl = Range("sqltbl")
Do Until sqltbl(i, 1) = Empty
    If sqltbl(i, -1) = "Y" Then
        strtyp = sqltbl(i, -2)
        Set usews = Worksheets(sqltbl(i, 1).Value)
        Set userng = usews.Range(sqltbl(i, 0).Value)
            'userng.Delete (xlShiftUp)
            userng.Clear
        
        Set userng = usews.Range(sqltbl(i, 0).Value)
        strquery = sqltbl(i, 2)
        
        strquery = sqltbl(i, 2)
        strquery = SetVblParameters(strquery)
    
        Set conn = GetConnection(Range("connstr"))
        Set rsdata = conn.Execute(strquery)
    
        If sqltbl(i, 2).Font.Bold = False Then
            If rsdata.BOF = True And rsdata.EOF = True Then
                Else
                   arr = rsdata.GetRows
                   c = UBound(arr, 1) + 1
                   r = UBound(arr, 2) + 1
                Set userng2 = Range(userng(1, 1), userng(r, c))
                arr = ArrayTranspose(arr)
                userng2.Value = arr
                For intColIndex = 4 To rsdata.Fields.Count - 1
                    userng(0, 1).Offset(0, intColIndex).Value = rsdata.Fields(intColIndex).Name
                Next
            End If
        End If
        sqltbl(i, 5).Value = "'" & Now()
'        NamedRanges userng, usews
'        sqltbl(i, 5).Value = "'" & Now()
    Else
    End If

i = i + 1
Loop

End Sub


Sub NamedRanges(userng As Range, usews As Worksheet)

TargetSheet = UCase(Replace(Left(usews.Name, 4), " ", ""))

j = 2
    Do While userng(j, 1) <> ""
        j = j + 1
Loop

LastRow = j - 1

k = 1

Do Until userng(0, k) = ""
    Range(userng(1, k), userng(LastRow, k)).Name = TargetSheet & Replace(userng(0, k), " ", "") & "Col"
    k = k + 1
Loop

End Sub


'References
Option Explicit
Option Compare Text

Function SetVblParameters(str As String)
Dim c, d As Integer
Dim ReplaceStr As String
Dim ReplaceStr1 As String

Do Until InStr(str, "{") = 0
c = InStr(str, "}")
d = InStr(str, "{")
ReplaceStr = Mid(str, d, c - d + 1)
ReplaceStr1 = Range(Mid(ReplaceStr, 2, Len(ReplaceStr) - 2))
str = Replace(str, ReplaceStr, "'" & ReplaceStr1 & "'")
Loop

SetVblParameters = str

End Function

Function GetConnection(connectionstring As String) As Connection
    Dim conn As New Connection
    
    On Error GoTo errh
    
    conn.CursorLocation = adUseClient
    conn.CommandTimeout = 150
    conn.Open connectionstring
    Set GetConnection = conn
    Exit Function
errh:
    Set GetConnection = Nothing
    MsgBox "Cannot connect to database at this time." & vbCrLf & Err.Description, , "Error"
    Exit Function
End Function

Public Function ArrayTranspose(InputArray)
'This function returns the transpose of
'the input array or range; it is designed
'to avoid the limitation on the number of
'array elements and type of array that the
'worksheet TRANSPOSE Function has.

'Declare the variables
Dim outputArrayTranspose As Variant, arr As Variant, p As Integer
Dim i As Long, j As Long, z As Long, msg As String

'Check to confirm that the input array
'is an array or multicell range
If IsArray(InputArray) Then

'If so, convert an input range to a
'true array
arr = InputArray

'Load the number of dimensions of
'the input array to a variable
On Error Resume Next

'Loop until an error occurs
i = 1
Do
z = UBound(arr, i)
i = i + 1
Loop While Err = 0

'Reset the error value for use with other procedures
Err = 0

'Return the number of dimensions
p = i - 2
End If

If Not IsArray(InputArray) Or p > 2 Then
msg = "#ERROR! The function accepts only multi-cell ranges and 1D or 2D arrays."
If TypeOf Application.Caller Is Range Then
ArrayTranspose = msg
Else
MsgBox msg, 16
End If
Exit Function
End If

'Load the output array from a one-
'dimensional input array
If p = 1 Then

Select Case TypeName(arr)
Case "Object()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Object
For i = LBound(outputArrayTranspose) To UBound(outputArrayTranspose)
Set outputArrayTranspose(i, LBound(outputArrayTranspose)) = arr(i)
Next
Case "Boolean()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Boolean
Case "Byte()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Byte
Case "Currency()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Currency
Case "Date()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Date
Case "Double()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Double
Case "Integer()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Integer
Case "Long()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Long
Case "Single()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Single
Case "String()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1)) As String
Case "Variant()"
ReDim outputArrayTranspose(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Variant
Case Else
msg = "#ERROR! Only built-in types of arrays are supported."
If TypeOf Application.Caller Is Range Then
ArrayTranspose = msg
Else
MsgBox msg, 16
End If
Exit Function
End Select
If TypeName(arr) <> "Object()" Then
For i = LBound(outputArrayTranspose) To UBound(outputArrayTranspose)
outputArrayTranspose(i, LBound(outputArrayTranspose)) = arr(i)
Next
End If

'Or load the output array from a two-
'dimensional input array or range
ElseIf p = 2 Then
Select Case TypeName(arr)
Case "Object()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Object
For i = LBound(outputArrayTranspose) To _
UBound(outputArrayTranspose)
For j = LBound(outputArrayTranspose, 2) To _
UBound(outputArrayTranspose, 2)
Set outputArrayTranspose(i, j) = arr(j, i)
Next
Next
Case "Boolean()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Boolean
Case "Byte()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Byte
Case "Currency()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Currency
Case "Date()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Date
Case "Double()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Double
Case "Integer()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Integer
Case "Long()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Long
Case "Single()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Single
Case "String()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As String
Case "Variant()"
ReDim outputArrayTranspose(LBound(arr, 2) To UBound(arr, 2), _
LBound(arr) To UBound(arr)) As Variant
Case Else
msg = "#ERROR! Only built-in types of arrays are supported."
If TypeOf Application.Caller Is Range Then
ArrayTranspose = msg
Else
MsgBox msg, 16
End If
Exit Function
End Select
If TypeName(arr) <> "Object()" Then
For i = LBound(outputArrayTranspose) To _
UBound(outputArrayTranspose)
For j = LBound(outputArrayTranspose, 2) To _
UBound(outputArrayTranspose, 2)
outputArrayTranspose(i, j) = arr(j, i)
Next
Next
End If
End If

'Return the transposed array
ArrayTranspose = outputArrayTranspose
End Function
