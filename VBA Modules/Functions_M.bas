Attribute VB_Name = "Functions_M"

Public Function DuplicateCountToScript(nums As Variant) As Scripting.Dictionary 'nums = 1D array
'1D array is turned into a dictionary with its Key as the contents of the array,
'and the item as the number of times the key occured in the original array, returns a dictionary
    Dim Dict As New Scripting.Dictionary
    For Each num In nums        'for each key in dictionary
        If Dict.Exists(num) Then    'if it exists, add count to dictionary
            Dict(num) = Dict(num) + 1
        Else
            Dict(num) = 1
        End If
    Next

   Set DuplicateCountToScript = Dict
End Function
Function ArrayRemoveDups(MyArray As Variant) As Variant 'MyArray= 1 '1D Arrays
'1D function takes 1D array and removes duplicates, returns as 1D array
    Dim nFirst As Long, nLast As Long, i As Long
    Dim item As String
    
    Dim arrTemp() As Variant
    Dim Coll As New Collection
 
    'Get First and Last Array Positions
    nFirst = LBound(MyArray)
    nLast = UBound(MyArray)
    ReDim arrTemp(nFirst To nLast)
 
    'Convert Array to String
    For i = nFirst To nLast
        arrTemp(i) = CStr(MyArray(i))
    Next i
    
    'Populate Temporary Collection
    On Error Resume Next
    For i = nFirst To nLast
        Coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.Clear
    On Error GoTo 0
 
    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)
    
    'Populate Array
    For i = nFirst To nLast
        arrTemp(i) = Coll(i - nFirst + 1)
    Next i
    
    'Output Array
    ArrayRemoveDups = arrTemp
 
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
        Arr = ArrayRemoveDups(Arr)  'if not empty, remove duplicates
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

Sub SortByDate(sheet_name As String, table_name As String, column_name As String)
'sorts table by column name on worksheet

Set book = ThisWorkbook
Set Sheet = book.Worksheets(sheet_name)
Dim tbl As ListObject
Set tbl = Sheet.ListObjects(table_name)
Dim sortcolumn As Range
Set sortcolumn = tbl.ListColumns("Date").DataBodyRange
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .Header = xlYes
   .Apply
End With
End Sub








