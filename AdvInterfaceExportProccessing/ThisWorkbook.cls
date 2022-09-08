VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const messageColumn = 14

Function getMsg(row As Long)
    
    Dim msg As String
    
    'find the column where text= "Message Text"
    '14 columns in output value. For now we'll just assume it's Row 1, Col 14
    
    msg = Cells(row, messageColumn).value
    
    getMsg = msg
End Function
Function getSegmentAtIteration(segName As String, msg As String, Optional iteration As Integer = 1)
    'gets the segment at the iteration. If not provided, get first iteration.
    Dim line As Variant
    
    Dim segment As String
    segment = "" 'default to blank
    
    'split msg.
    Dim msgSplit() As String
    msg = Replace(msg, vbCrLf, vbCr)
    msg = Replace(msg, vbLf, vbCr)
    msgSplit = Split(msg, vbCr)
    
    'loop through the array
    For Each line In msgSplit
        If Left(line, 3) = segName Then
            'get iteration of line
            Dim sigIter As String
            sigIter = getSegmentValueAtIndex(line, 1)
            If sigIter = iteration Then
                segment = line
            End If
            
        End If
    Next
    
    
    getSegmentAtIteration = segment
End Function

Function getSegmentValueAtIndex(segment As Variant, index As Variant)
    'gets the value of the segment at a given index
    '0 returns the Segment name
    'MSH is handled differently. if it's anything other than 1, return as normal. 1 returns a 1 to match the other segments.
    
    Dim line As Variant
    'split msg.
    
    Dim value As String
    
     'split
    Dim segSplit() As String
    segSplit = Split(segment, "|")
    
    'check for MSH
    If Left(segment, 3) <> "MSH" Then
        
        value = segSplit(index)
    Else
        'msh
        If index = 0 Then
            value = segSplit(index)
        ElseIf index = 1 Then
            'msh.1
            value = "1"
        ElseIf index > 1 Then
             value = segSplit(index - 1)
        End If
    
    End If
    
    getSegmentValueAtIndex = value
End Function

Function countSegments(segName As String, msg As String)
    'gets count of segments with name
    Dim count, i As Integer
    
    
    'split the message on \r\n into an array
    Dim msgSplit() As String
    msg = Replace(msg, vbCrLf, vbCr)
    msg = Replace(msg, vbLf, vbCr)
    msgSplit = Split(msg, vbCr)
    
    'loop through array, checking the first 3 chars of each one
    For Each line In msgSplit
        If Left(line, 3) = segName Then
            count = count + 1
        End If
        
        
    Next
    
    
    countSegments = count
    
End Function



Function getSegmentNames()
    'get the list of segment names in sheet.
    'unique list.
    Dim i As Long
    Dim msg As String
    Dim segName As String
    Dim segment, val As Variant
    Dim isInArray As Boolean
    isInArray = False
    
    Dim sizeOfSegmentNames As Integer
    Dim pointToAddNewValue As Integer
    sizeOfSegmentNames = 0
    pointToAddNewValue = 0
    
    Dim segmentNames() As String
    'do initial redim
    ReDim segmentNames(sizeOfSegmentNames)
    
    
    'loop through each message in sheet, and add new segments to array.
    Dim rows As Long
    rows = getRowsInSheet()
    
    'start with row 2.
    For i = 2 To rows
        msg = getMsg(i)
        
        'split msg to segments
        Dim msgSplit() As String
        msg = Replace(msg, vbCrLf, vbCr)
        msg = Replace(msg, vbLf, vbCr)
        msgSplit = Split(msg, vbCr)
        
        'get first index of each segment
        For Each segment In msgSplit
            segName = getSegmentValueAtIndex(segment, 0)
            'loop throuhgh segmentNames and determine if it's in the array. if not, redim and add it.
            For Each val In segmentNames
                If val = segName Then
                    isInArray = True
                End If
            Next val
            
            If isInArray = False Then
                                
                If sizeOfSegmentNames <> 0 Then
                    ReDim Preserve segmentNames(sizeOfSegmentNames)
                End If
                
                
                
                segmentNames(pointToAddNewValue) = segName
                pointToAddNewValue = pointToAddNewValue + 1
                sizeOfSegmentNames = sizeOfSegmentNames + 1

            End If
            
        Next segment
        
    Next i
    
    getSegmentNames = segmentNames
End Function



Function getMessageValuesAtSegmentAndIndex(msg As String, segmentName As String, index As Variant)
    'returns a comma separated value list of all the values for all segments with the given name, at the specified index. Leaves the final comma on the list.
    
    'get count of segments with the seg name
    Dim output As String
    Dim count, i As Integer
    Dim segment As String
    count = countSegments(segmentName, msg)
    
    For i = 1 To count
        segment = getSegmentAtIteration(segmentName, msg, i)
        'get the value at index.
        output = output + getSegmentValueAtIndex(segment, index) + ","
        
    Next i
     getMessageValuesAtSegmentAndIndex = output
End Function

Function getRowsInSheet()
    'assuming flat results as exported by PK.
    Dim count As Long
    
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets(1)
    Set rng = ws.UsedRange
    
    count = rng.rows.count
    
    getRowsInSheet = count
End Function



Sub searchMessages(segmentName As String, index As Integer, searchVal As String)
    'loop through all values of messages.
    'if the value is in the field, highlight cell yellow
    
    'clear any existing highlighting
    resetHighlighting
    
    Dim msg As String
    Dim rows, i As Long
    Dim value As String
    Dim output As String
    Dim count, j As Integer
    Dim seg As String
    

    
    rows = getRowsInSheet()
    
    For i = 2 To rows
        msg = getMsg(i)
        
        'get number of segments with segment name
        count = countSegments(segmentName, msg)
        
        For j = 1 To count
            segment = getSegmentAtIteration(segmentName, msg, j)
            'get the value at index.
            value = getSegmentValueAtIndex(segment, index)
            
            If InStr(value, searchVal) <> 0 Then
                'highlight at i
                highlightMsg (i)
            End If
            
            
        Next j
    Next i
    
End Sub

Sub highlightMsg(row As Integer)
    'msg column is
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets(1)
    Set rng = ws.rows(row)
    rng.Interior.Color = vbYellow
    
    
End Sub

Sub resetHighlighting()
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Interior.Color = xlNone
End Sub

Function getSheetValuesAtSegmentAndIndex(segment As String, field As Integer)
   
    Dim msg As String
    Dim rows, i As Long
    Dim values As String
    
    rows = getRowsInSheet()
    
    For i = 2 To rows
        msg = getMsg(i)
        values = values + getMessageValuesAtSegmentAndIndex(msg, segment, field)
    
    Next i
    
    'remove duplicates from list
    values = deDuplicateList(values)
    
    'remove final comma
    values = Left(values, Len(values) - 1)
    
    'alphabetize
    values = alphabetizeList(values)
    
    'replace commas with CRLF
    values = Replace(values, ",", vbCrLf)
    
    
    getSheetValuesAtSegmentAndIndex = values
End Function
Function alphabetizeList(list As Variant)
    'calls the quicksort algo.
    
    'make list
    Dim listArr() As String
    listArr = Split(list, ",")
    Call QuickSort(listArr, 0, UBound(listArr))
    
    'turn back to list
    Dim output As String
    Dim tempVal As Variant
    For Each tempVal In listArr
        output = output + tempVal + ","
    Next tempVal
        
    alphabetizeList = output
End Function

Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub

Function deDuplicateList(list As Variant)
    Dim tempArray() As String
    Dim splitArray() As String
    Dim i, j As Integer
    Dim sizeOfTempArray As Integer
    Dim pointToAddValue As Integer
    Dim isInArray As Boolean
    
    isInArray = False
    
    sizeOfTempArray = 0
    pointToAddValue = 0
    
    Dim splitVal, tempVal As Variant
    
    'split the list on comma
    splitArray = Split(list, ",")
    
    'initial redim
    ReDim tempArray(sizeOfTempArray)
    
    'loop through each element in arr, and search through the temp array for it's existance.
    For Each splitVal In splitArray
        For Each tempVal In tempArray
            If splitVal = tempVal Then
                isInArray = True
            End If
        Next tempVal
        
        If isInArray = False Then
            If sizeOfTempArray <> 0 Then
                    ReDim Preserve tempArray(sizeOfTempArray)
                End If
                
                
                
                tempArray(pointToAddValue) = splitVal
                pointToAddValue = pointToAddValue + 1
                sizeOfTempArray = sizeOfTempArray + 1
         End If
            
        
    Next splitVal
    
    'turn back to a list
    Dim output As String
    For Each tempVal In tempArray
        output = output + tempVal + ","
    Next tempVal

    deDuplicateList = output
 
End Function


Sub main()
Attribute main.VB_Description = "Runs the macro"
Attribute main.VB_ProcData.VB_Invoke_Func = "G\n14"
    

    'load form
    Form.Show (False)

End Sub



