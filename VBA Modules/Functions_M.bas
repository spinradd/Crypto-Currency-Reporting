Attribute VB_Name = "Functions_M"

Function Remove_duplicates_array(arr_with_dups As Variant) As Variant 'arr_with_dups= 1 '1D Arrays
'1D function takes 1D array and removes duplicates, returns as 1D array
    
    
    Dim first_item As Long, last_item As Long, i As Long
    Dim item As String
    
    Dim intermediate_arr() As Variant
    Dim coll As New Collection
 
    'find length of current array; resize intermediary array to fit
    first_item = LBound(arr_with_dups)
    last_item = UBound(arr_with_dups)
    ReDim intermediate_arr(first_item To last_item)
 
    'normalize to string type
    For i = first_item To last_item
        intermediate_arr(i) = CStr(arr_with_dups(i))
    Next i
    
    'add items to collection
    On Error Resume Next
    For i = first_item To last_item
        coll.Add intermediate_arr(i), intermediate_arr(i)
    Next i
    Err.Clear
    On Error GoTo 0
 
    'resize new array
    last_item = coll.Count + first_item - 1
    ReDim intermediate_arr(first_item To last_item)
    
    'for new array
    For i = first_item To last_item
        intermediate_arr(i) = coll(i - first_item + 1)
    Next i
    
    Remove_duplicates_array = intermediate_arr
 
End Function


Function BlankRemover(ArrayToCondense As Variant) As Variant() '1D Arrays
'takes 1D arrays and removesblanks, returns as 1D array

Dim ArrayWithoutBlanks() As Variant
Dim CellsInArray As Variant
ReDim ArrayWithoutBlanks(0 To 0) As Variant

IsAllBlank = True
For Each CellsInArray In ArrayToCondense
    If CellsInArray <> "" Then
        IsAllBlank = False
        ArrayWithoutBlanks(UBound(ArrayWithoutBlanks)) = CellsInArray
        ReDim Preserve ArrayWithoutBlanks(0 To UBound(ArrayWithoutBlanks) + 1)
    End If
    
    'MsgBox Join(ArrayWithoutBlanks, vbCrLf)

Next CellsInArray


'get rid of extra blank space
If IsAllBank = False Then
        Do While ArrayWithoutBlanks(UBound(ArrayWithoutBlanks)) = Empty
            ReDim Preserve ArrayWithoutBlanks(0 To UBound(ArrayWithoutBlanks) - 1)
        Loop
        BlankRemover = ArrayWithoutBlanks
    
    Else
    
        ReDim ArrayWithoutBlanks(0 To 0)
        BlankRemover = ArrayWithoutBlanks
End If
End Function

Public Function DoesArrayExist(Arr() As Variant) As Variant

Dim ArrayResults(0 To 1) As Variant

DoesArrExist = True
IsArrEmpty = False
Dim NoArr() As Variant

If (Not Not Arr) = 0 Then       'if array never intitialized then it doesn't exist,
        DoesArrExist = False
        IsArrEmpty = True        'if list doesn't exist, then it's empty (by default)
        ReDim Preserve NoArr(0)
        NoArr(0) = "Does Not Exist"
        DoesArrayExist = NoArr
    Else
        For Each item_in_array In Arr       'test if array is empty. If there is one non-blank cell, change bool value
            If item_in_array <> Empty Then
                IsArrEmpty = False
            End If
        Next
End If

If DoesArrExist = True Then            'if array isn't empty, remove blanks, remove dups
        Arr = BlankRemover(Arr)     'if not empty, remove blanks
        Arr = Remove_duplicates_array(Arr)  'if not empty, remove duplicates
        DoesArrayExist = Arr
ElseIf DoesArrExist = False Then    'if array is empty,  go to next column
        ReDim Preserve NoArr(0)
        NoArr(0) = "No Data found"
        DoesArrayExist = NoArr
End If

End Function


Public Function GetArray(SheetName As String, ListObject_name As String, _
                                    column_name As String)
'function needs sheetname, listobject name, column header title
'will create a 1D array from table data body range, filled with blanks and duplicate values
    Dim arrbase() As Variant
    entry_count = 0
    For Each cell In Worksheets(SheetName).ListObjects(ListObject_name).ListColumns(column_name).DataBodyRange
    
                              'if first entry, redim to hold one spot "(0)"
                           If entry_count = 0 Then
                               ReDim Preserve arrbase(0)
                               arrbase(0) = cell
                            'For all subsequent entries extend array by 1 and enter contents in cell
                           Else
                               ReDim Preserve arrbase(UBound(arrbase) + 1)
                               arrbase(UBound(arrbase)) = cell
                           End If
                           entry_count = entry_count + 1
    Next cell
    
   ' MsgBox Join(arrbase, vbCrLf)
    GetArray = arrbase

End Function

Sub SortWorksheetsAlphabetially()


'Turn off screen updating
Application.ScreenUpdating = False

'Create variables
Dim book As Workbook
Dim wsht_count As Integer
Dim i As Integer
Dim j As Integer
Set book = Workbooks(ActiveWorkbook.Name)

'Count the number of worksheets
wsCount = book.Worksheets.Count

'Loop through all worksheets and move
For i = 1 To wsht_count - 1
    For j = i + 1 To wsht_count
        If book.Worksheets(j).Name < book.Worksheets(i).Name Then
            book.Worksheets(j).Move before:=book.Worksheets(i)
        End If
    Next j
Next i

'Turn on screen updating
Application.ScreenUpdating = True

End Sub

Sub SortByCol(sheet_name As String, table_name As String, column_name As String)
'sorts table by column name on worksheet

Set book = ThisWorkbook
Set Sheet = book.Worksheets(sheet_name)
Dim tbl As ListObject
Set tbl = Sheet.ListObjects(table_name)
Dim sortcolumn As Range
Set sortcolumn = tbl.ListColumns(column_name).DataBodyRange
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .Header = xlYes
   .Apply
End With
End Sub








