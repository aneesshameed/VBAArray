VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''/'''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''Code Starts''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''/'''''''''''''''''''''''''''''''''''''''''

'Public variable of the class
Private LocalArray() As Variant


'Initialization of array

Private Sub Class_Initialize()
    
    'We get an array with size 1. This is very important, If we dim a dynamic array then before using array, it need to
    'be redimed to a size of 1
    ReDim LocalArray(0)
    
End Sub


'Pushes an item to array. New item will go to the last index of array
'Logic: For each item in argument items() increase size of array and add item to last index

Public Sub Push(ParamArray items() As Variant)
    
    'Loop for all items()
    For i = 0 To UBound(items)
    
        'Redim the LocalArray() and push each item to new index of LocalArray()
        ReDim Preserve LocalArray(0 To GetReDimSize(LocalArray))
        LocalArray(GetUbound(LocalArray)) = items(i)
    Next i
    
End Sub


'Removes an item from the end of array
'Logic: To remove item from end, ReDim by setting ubound on less

Public Function Pop()

    'ReDim array from 0 to its UBound - 1
    ReDim Preserve LocalArray(0 To UBound(LocalArray) - 1)
    
End Function


'Removes an item from the beginning of array
'Logic: Create a temp(), parse through index 1 to Ubound of LocalArray(), and add each item to temp()
'At the end clone temp() to LocalArray()

Public Sub Shift()
    
    'Initialize temp()
    Dim temp As New VBAArray

    'Skip item in index 0 of LocalArray() and add all others items to temp
    For i = 1 To GetUbound(LocalArray)

        'push items to temp()
        temp.Push (LocalArray(i))
    Next i

    'Now clone temp() to LocalArray()
    LocalArray = temp.Arrayify
End Sub

'Adds items to the beginning of array
'Logic: Create a new temp() array, add new item to temp()'s 0th index onwards. Parse through the LocalArray() and push each item
'to temp()
'At the end clone LocalArray() with newly created temp()

Public Sub UnShift(ParamArray items() As Variant)
    
    'Create a temp()
    Dim temp As New VBAArray
    
    'For all items in array argument items(), push each to temp
    For i = 0 To UBound(items)
        
        'push item to temp()
        temp.Push items(i)
    Next i
    
    'Now push all items in LocalArray() to temp()
    For i = 0 To GetUbound(LocalArray)
        
        'push item to temp()
        temp.Push LocalArray(i)
    Next i
   
   'Now clone temp() to LocalArray()
    LocalArray = temp.Arrayify
    
End Sub


'Creates a new VBAArray by selecting items specified by StartIndex and EndIndex
'Return type will be VBAArray class and hence the variable accepting return should be of same type to store it.
'Logic: Create a temp() and move items in LocalArray() if index is between StartIndex and EndIndex
'Return Slice.

Public Function Slice(Optional StartIndex As Variant, Optional EndIndex As Variant) As VBAArray
    
    On Error GoTo SendNullArray
    'Initialize temp
    Dim temp As New VBAArray
    
    'If StartIndex is missing then user need all items in old array to be pushed to new array, hence StartIndex will be 0
    If IsMissing(StartIndex) Then
        StartIndex = 0
    Else
        
        'If StartIndex is give as negative, then operation starts from end of array
        If StartIndex < 0 Then
            StartIndex = GetUbound(LocalArray) + StartIndex
        End If
    End If
    
    
    'If EndIndex is missing then user needs to get all items after StartIndex
    If IsMissing(EndIndex) Then
        EndIndex = GetUbound(LocalArray)
    Else
        
        'If EndIndex is negative number then operation is supposed to start from end of array
        If EndIndex < 0 Then
            EndIndex = GetUbound(LocalArray) + EndIndex
        Else
            EndIndex = WorksheetFunction.Min(EndIndex - 1, GetUbound(LocalArray))
        End If
    End If
    
    
    'Push items to temp if index is between the Start and Start+HowMany
    For i = StartIndex To EndIndex
        temp.Push LocalArray(i)
    Next i
    
    'Now clone temp and return it.
    Set Slice = temp
    
    'And exit function
    Exit Function
    
    'If any error has happened then return a null array
SendNullArray:
    Set Slice = New VBAArray
    
End Function


'Add or remove items from an array
'Logic: Create a temp() array. Push all items of LocalArray from index 0 to StartIndex to temp().
'NewItems is an optional variable, If it is present then those need to be pushed to temp()
'If HowMany is not zero then there are some items that need to be removed from LocalArray.
'Hence take items that are from StartIndex+HowMany to ubound of LocalArray and push all to temp()
'Clone temp() to LocalArray

Public Sub Splice(ByVal StartIndex As Integer, Optional HowMany As Variant, Optional NewItems As Variant)
    
    On Error GoTo SendFullArray
    
    'Initialize temp
    Dim temp As New VBAArray
    
    'If start index is negative then operation is supposed to start from end of array
    If StartIndex < 0 Then
        StartIndex = UBound(LocalArray) + StartIndex
    Else
        
        'Else if user is giving a StartIndex which is greater than UBound of array, then take UBound as StartIndex.
        'In such case there will not be any items removed from LocalArrat
        StartIndex = WorksheetFunction.Min(StartIndex, UBound(LocalArray) + 1)
    End If
    
    'If StartIndex is greater than zero then push all items from 0 to StartIndex into temp()
    If StartIndex > 0 Then
       For i = 0 To StartIndex - 1
            temp.Push LocalArray(i)
       Next i
    End If
    
    'Now push NewItems to temp()
    If Not IsMissing(NewItems) Then
    
        'If NewItems is an array, then push all to temp()
        If IsArray(NewItems) Then
             For i = 0 To UBound(NewItems)
                    temp.Push NewItems(i)
                Next i
        Else
        
            'Else, it will be single item and hence push the same.
            temp.Push NewItems
        End If
    End If
    
    'Now push items between StartIndex+HowMany and UBound of LocalArray to temp()
    If Not IsMissing(HowMany) And IsNumeric(HowMany) Then
        For i = StartIndex + HowMany To UBound(LocalArray)
            temp.Push LocalArray(i)
        Next i
    End If
    
    'Clone temp() to LocalArray()
    LocalArray = temp.Arrayify
    
SendFullArray:
    
End Sub

'Reverse an array
'Note: There is no reverse iteration in VBA
'Logic: Start from 0 to UBound of LocalArray, push each items from the end of array to temp()
'Clone temp() to LocalArray

Public Sub Reverse()
    
    'Initialize temp()
    Dim temp As New VBAArray
    
    'Iterate from 0 to UBound
    For i = 0 To UBound(LocalArray)
        
        'Push items from end of array
        temp.Push (LocalArray(UBound(LocalArray) - i))
    Next i
    
    LocalArray = temp.Arrayify
End Sub


'Joins array seperated by the Delimiter passed
'Logic: Create temp as string.
'Starting from 0 to UBound concatenate values from LocalArray and delimiter to temp

Public Function Join(Optional Delimiter As Variant)
    
    'Initialize temp as string
    Dim temp As String
    temp = ""
    
    'If delimiter is missing then set Delimiter as ','
    If IsMissing(Delimiter) Then
        Delimiter = ","
    End If
    
    'Loop for items from 0 to Ubound - 1
    For i = 0 To UBound(LocalArray) - 1
        temp = temp & LocalArray(i) & Delimiter
    Next i
    
    'And then concatenate value in last index
    temp = temp & LocalArray(i)
    
    'Clone temp() to LocalArray()
    Join = temp
End Function


'Returns item specified by the index.
'If a NewValue has been passed then index will be set to NewValue and the same will be passed

Public Function Value(ByVal Index As Variant, Optional NewValue As Variant)
    
    'If index is greater then Ubound then return undefined
    If Index > UBound(LocalArray) Then
        Value = "Undefined"
    Else
        If Not IsMissing(NewValue) Then
            
            'If NewValue is passed to
            LocalArray(Index) = NewValue
            Value = LocalArray(Index)
        Else
        'Else return value
            Value = LocalArray(Index)
        End If
    End If
End Function

'Sends all values in VBAArray as proper array

Public Function Arrayify()
    
    'Initialize temp
    Dim temp() As Variant
    
    'Push everything to temp
    For i = 0 To UBound(LocalArray)
        ReDim Preserve temp(0 To GetReDimSize(temp))
        temp(UBound(temp)) = LocalArray(i)
    Next i
    
    'Send back temp
    Arrayify = temp
    
End Function


'Returns length of an array

Public Function Length()
    
    On Error GoTo Err:
        'Length us Ubound
        If LocalArray(0) = Empty Then
            Length = 0
        Else
            Length = UBound(LocalArray) + 1
        End If
        Exit Function
Err:
        Length = 0
End Function


'Returns index of an item from array
'Logic: If item is present in LocalArray() then match will return position of item.
'Else vba will throw error.

Public Function IndexOf(item As Variant)

    'Attached error catching.
    On Error GoTo Err:
        
        'Get index by match function and exit function
        IndexOf = Application.Match(item, LocalArray, False) - 1
        Exit Function
        
        'In case of error return -1
Err:
        IndexOf = -1
End Function

'''''''''''''''''''''''''''''''''''''''''/'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''Private Functions'''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''/'''''''''''''''''''''''''''''''''''''''''
Private Function GetReDimSize(Element)
    
    'If first index is empty then return 0 else return the size of array
    On Error GoTo Err
    If Element(UBound(Element)) = Empty Then
        GetReDimSize = 0
    Else
        GetReDimSize = UBound(Element) + 1
    End If
    Exit Function
Err:
    GetReDimSize = 0
End Function

Private Function GetUbound(Element)

    'Return size of array, and on error return 0
    On Error GoTo Err
    GetUbound = UBound(Element)
    Exit Function
Err:
    GetUbound = -1
End Function

