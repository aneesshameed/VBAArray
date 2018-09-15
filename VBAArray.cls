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
    
    'We get an with size 1. This is very important, If we dim a dynamic array then before using array, it need to
    'be redimed to a size of 1
    ReDim LocalArray(0)
    
End Sub


'Pushes an item to array. New item will go to the last index of array
'Logic: For each items in Arguments() increase size of array and add item to last index


Public Sub Push(ParamArray items() As Variant)
    
    'Redim array, extent its ubound by one, and push item to newly created index
    For I = 0 To UBound(items) 'Cannot run GetUbound on items as it is not an Array but ParamArray
    
        'Redim the LocalArray() and push each item to new index of LocalArray()
        ReDim Preserve LocalArray(0 To GetReDimSize(LocalArray))
        LocalArray(GetUbound(LocalArray)) = items(I)
    Next I
    
End Sub


'Add an item to the beginning of array
'Logic: Create a new temp() array, add new item to temp()'s 0th index onwards. Parse through the LocalArray() and push each item
'to temp()'s next index.
'At the end clone LocalArray() with newly created temp()


Public Sub UnShift(ParamArray items() As Variant)
    
    'Create a temp() and redim its size
    Dim temp() As Variant
    
    'For all elements in array Arguments() add each to temp
    For I = 0 To UBound(items) 'Cannot run GetUbound on items as it is not an Array but ParamArray
        ReDim Preserve temp(0 To GetReDimSize(temp))
        temp(GetUbound(temp)) = items(I)
    Next I
    
    'Push all elments in LocalArray() to temp()
    For I = 0 To GetUbound(LocalArray)
    
        'Redim the temp() and push each item to new index of temp()
        ReDim Preserve temp(0 To GetReDimSize(temp))
        temp(GetUbound(temp)) = LocalArray(I)
    Next I
   
   'Now clone temp() to LocalArray()
    LocalArray = temp
    
End Sub


'Removes an item from the beginning of array
'Logic: Create a temp(), parse through index of LocalArray() index 1 to last index and add each item to temp()
'At the end clone temp() to LocalArray()


Public Sub Shift()
    
    'Initialize temp()
    Dim temp() As Variant
    ReDim temp(0)
        
    'Skip item in index 0 of LocalArray() and add all others items to temp
    For I = 1 To GetUbound(LocalArray)
        
        'Redim the array and push each item to new index of temp()
        ReDim Preserve temp(0 To GetReDimSize(temp))
        temp(GetUbound(temp)) = LocalArray(I)
    Next I
    
    'Now clone temp() to LocalArray()
    LocalArray = temp
End Sub


'Creates a new VBAArray class by selecting elements specified by Start index and count
'Return type will be VBAArray class and hence the variable accepting return should be ready to store it.
'Logic: Create a temp() and move LocalArray() items specified by arguments to temp()
'Create a new VBAArray class with name Slice and push items in temp() to new class created.
'Return class to parent.


Public Function Slice(Optional StartIndex As Variant, Optional EndIndex As Variant) As VBAArray
    
    On Error GoTo SendNullArray
    'Initialize temp()
    Dim temp() As Variant
    
    
    'If StartIndex is missing then user need all items in old array to a new array
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
            EndIndex = WorksheetFunction.Min(StartIndex + EndIndex, GetUbound(LocalArray))
        End If
    End If
    
    
    'Push items to temp() if index is between the Start and Start+HowMany
    For I = 0 To GetUbound(LocalArray)
        If I >= StartIndex And I <= EndIndex Then
            ReDim Preserve temp(0 To GetReDimSize(temp))
            temp(GetUbound(temp)) = LocalArray(I)
        End If
    Next I
    
    'Now clone temp() to LocalArray(). For that create a new class of VBAArray and name it Slice.
    'Move all the items in temp() to Slice.
    Set Slice = New VBAArray
    For I = 0 To GetUbound(temp)
        Slice.Push (temp(I))
    Next I
    
    'And exit function
    Exit Function
    
SendNullArray:
    Set Slice = New VBAArray
    
End Function


'Add or remove items from an array
'Logic: If argument HowMany is not zero then there are some items that need to be removed from array.
'NewItems is an optional variable, If it is present then those need to be pushed to the index from Start


Public Sub Splice(ByVal StartIndex As Integer, Optional HowMany As Variant, Optional NewItems As Variant)
    
    On Error GoTo SendFullArray
    
    
    'Initialize temp()
    Dim temp() As Variant
    
    
    'If start index is negative then operation is supposed to start from end of array
    If StartIndex < 0 Then
        StartIndex = GetUbound(LocalArray) + StartIndex + 1
    Else
        
        
        'Else if user is giving a StartIndex which is greater than UBound of array, then take UBound as StartIndex
        StartIndex = WorksheetFunction.Min(StartIndex, GetUbound(LocalArray))
    End If
        
        
    'If HowMany is being defined then we need to take out items which are between StartIndex and HowMany
    If Not IsMissing(HowMany) And IsNumeric(HowMany) Then
        
        HowMany = WorksheetFunction.Min(HowMany, GetUbound(LocalArray))
        
        
        'For all indexes in LocalArray
        For I = 0 To GetUbound(LocalArray)
            
            
            'If index is less than StartIndex and greater then StartIndex+HowMany-1 then we need to push those items to temp()
            If I < StartIndex Or I >= StartIndex + HowMany Then
                ReDim Preserve temp(0 To GetReDimSize(temp))
                temp(GetUbound(temp)) = LocalArray(I)
            End If
        Next I
        
    Else
    
        StartIndex = StartIndex - 1
        
        
        'If there are nothing to be deleted then all items from LocalArray() need to be pushed to temp() till one less than StartIndex
        For I = 0 To StartIndex
            ReDim Preserve temp(0 To GetReDimSize(temp))
            temp(GetUbound(temp)) = LocalArray(I)
        Next I
    End If
    'Now we new have temp() that contains only items that are needed
    
    
    'Now add NewItems to temp(), these items needs to be pushed after the StartIndex
    If Not IsMissing(NewItems) Then
        
        
        'If NewItems is passed as an array, then push all items in NewItems to temp(). Otherwise NewItems is just a string, just push
        'only one item to array.
        If IsArray(NewItems) Then
            
            
            'Index of last non empty cell, Soring this value as we are soon going to ReDim temp()
            Dim UboundAfterRemoval As Integer
            UboundAfterRemoval = GetUbound(temp)
            
            
            'Redim array to length spanning earlier length and length of NewItems()
            If UboundAfterRemoval = -1 Then
                
                
                'If UboundAfterRemoval is less than zero then it means user have asked to take nothing from LocalArray()
                'In this case temp() need to be ReDimmed to 0 to UBound(NewItems)
                'Consider NewItems() as [8,9,10], Ubound(NewItems) will be 2
                'Hence do a ReDim of array to 0 to (2) and push NewItems
                ReDim Preserve temp(0 To GetUbound(NewItems))
                For I = 0 To GetUbound(NewItems)
                    ReDim Preserve temp(0 To GetReDimSize(temp))
                    temp(GetUbound(temp)) = NewItems(I)
                Next I
                LocalArray = temp
                Exit Sub
            ElseIf UboundAfterRemoval = StartIndex - 1 Then
                
                
                'Consider Current temp() as [1,2,3,4,5], User asked to splice 3 items from index 2.
                'This will lead to [1,2] and UboundAfterRemoval as 1.
                'Hence UboundAfterRemoval becomes equals to  StartIndex
                'This will lead to situation that need to push all items to the end of array
                ReDim Preserve temp(0 To GetUbound(temp) + UBound(NewItems) + 1)
                
                For I = 0 To GetUbound(NewItems)
                    temp(StartIndex + I) = NewItems(I)
                Next I
                LocalArray = temp
                Exit Sub
            Else
                
                
                'Consider Current temp() as [1,2,3,4,5], User asked to splice from 2 items from index 1.
                'This will lead to [1,4,5] and UboundAfterRemoval as 2. StartIndex is 1 in this case.
                'Means, UboundAfterRemoval is greater than StartIndex.
                'This implies all items that are from StartIndex to UboundAfterRemoval is supposed to be moved towards the end of temp()
                
                
                'We first need to ReDim temp() to accomodate all items in NewItems()
                ReDim Preserve temp(0 To GetUbound(temp) + UBound(NewItems) + 1)
                
                
                'Now move items to end of array
                Dim J As Integer
                J = 0
                For I = StartIndex To UboundAfterRemoval
                    temp(UBound(temp) - J) = temp(UboundAfterRemoval - J)
                    J = J + 1
                Next I
                
                
                'Now update index starting from StartIndex with new items
                For I = 0 To GetUbound(NewItems)
                    temp(StartIndex + I) = NewItems(I)
                Next I
                LocalArray = temp
                Exit Sub
            End If
        Else
            
            
            'Else item is a string. Increase temp() by just 1. Shift items by one index up.
            ReDim Preserve temp(0 To UBound(temp) + 1)
            
            
            'Get Last index of temp()
            UboundAfterRemoval = UBound(temp)

            Do While UboundAfterRemoval > StartIndex
                
                
                'Update nth index of temp() by n-length of NewItems() of temp. This way we are moving items to the end of temp()
                temp(UboundAfterRemoval) = temp(UboundAfterRemoval - 1)
                UboundAfterRemoval = UboundAfterRemoval - 1
            Loop
             
            
            'And add NewItems to index next to StartIndex
            temp(UboundAfterRemoval) = NewItems
            LocalArray = temp
            Exit Sub
        End If
    Else
        LocalArray = temp
    End If
    Exit Sub
SendFullArray:
    
End Sub

'Removes an item from the end of array
'Logic: To remove item from end, ReDim by setting ubound on less
Public Function Pop()

    'ReDim array from 0 to its UBound - 1
    ReDim Preserve LocalArray(0 To UBound(LocalArray) - 1)
End Function

'Removes an item from array specified by index
Public Sub Remove(ByVal Index As Integer)
    
    'Initialize temp()
    Dim temp() As Variant
    ReDim temp(0)
    
    'Loop for all items in LocalArray()
    For I = 0 To UBound(LocalArray)
    
        'If I is not equal to the index then push item to temp()
        If I <> Index Then
            ReDim Preserve temp(0 To GetReDimSize(temp))
            temp(UBound(temp)) = LocalArray(I)
        End If
    Next I
    
    'Clone temp() to LocalArray()
    LocalArray = temp
End Sub


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

'Send all values in VBAArray as proper array
Public Function Arrayify()
    
    'Initialize temp
    Dim temp() As Variant
    
    'Push everything to temp
    For I = 0 To UBound(LocalArray)
        ReDim Preserve temp(0 To GetReDimSize(temp))
        temp(UBound(temp)) = LocalArray(I)
    Next I
    
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


