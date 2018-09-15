# Helper class for VBA arrays

Working on arrays with vba is bit difficult as there is no proper array methods for array manipulation. Use this vba class to easily work on vba array.

To get started import VBAArray.cls file to your project and initialize a new class by
```
 Dim ArrayName as New VBAArray
```
For the rest of this document lets consider an array of fruits
```
Dim Fruits as new VBAArray
```
## Array Methods

### Add items to an array
```
Fruits.Push "Apple"
Fruits.Push "Banana", "Grape", "Date"
```
Result array class will be having "Apple", "Banana", "Grape", "Date"

### Remove items from an array

```
Fruits.Pop
```
Result array class will be having "Apple", "Banana", "Grape"

### Add items to the beginning of an array 

```
Fruits.UnShift "Orange", "Mango"
```
Result array class will be having "Orange", "Mango", "Apple", "Banana", "Grape"

### Remove item from the beginning of array
```
Fruits.Shift
````
Result array will be "Mango", "Apple", "Banana", "Grape". "Orange" will be removed from array class

### Get value of an index
```
Dim ThisFruit as String
ThisFruit = Fruits.Value(0)
```
Result will be "Mango"

### Set value of an index
```
Dim ThisFruit as String
ThisFruit = Fruits.Value(0, "New Mango")
```
Result will be "New Mango"

### Get index of an item
```
Dim Pos as Integer
Pos = Fruits.IndexOf("Mango")
```
Result in Pos will be 0

### Get length of an array
```
Dim Length as Integer
Length = Fruits.Length
```
Result in Length will be 5.

### Get all values in Fruits to an array.
```
Dim FruitsAsArray() As Variant
FruitsAsArray = Fruits.Arrayify
Debug.Print FruitsAsArray(0)
```
FruitsAsArray will be an array and not a class. You can write the result to sheet as follows

Writes data in columns
```
Sheet1.Range("A1").resize(1, UBound(FruitsAsArray)).Value = FruitsAsArray
```

Writes data in rows
```
Sheet1.Range("A1").Resize(UBound(FruitsAsArray), 1).value = Application.WorksheetFunction.Transpose(FruitsAsArray)
```

### Slice an array

Returns a new VBAArray class with items from current array and indexes specified. Return type is also a class of VBAArray and hence variable that accept the result should be a type VBAArray class. Use negative numbers to specify positions from the end of array.

```
Dim NewFruits As New VBAArray
Set NewFruits = Fruits.Slice(0, 3)
````
Result NewFruits will be a class of VBAArray. NewFruits will be having items of "Mango", "Apple", "Banana"

### Splice an array

Helps to add or remove items from array. Use negative numbers to specify positions from the end of array.

For all splice example we will consider Fruits array as "Mango", "Apple", "Banana", "Grape", "Orange". 

* Remove an item from Fruits
```
Fruits.Splice 1, 1
```
Index 1, "Apple" will be removed and Fruits array will be "Mango", "Banana", "Grape","Orange".

* Remove more than one item from Fruits
```
Fruits.Splice 1, 2
```
Including index 1, two items will be removed. That is "Apple", "Banana" will be removed. Fruits array will be "Mango", "Grape","Orange".

* Don't remove any item, but add an item to Fruits.
```
Fruits.Splice 2, 0, "Lemon"
```
New items will be added after index 1, Fruits array will be now having "Mango", "Apple", "Banana","Lemon", "Kiwi",  "Grape", "Orange".

* Remove two items and add other three onto Fruits
```
Fruits.Splice 2, 2 Array("Watermelon", "Pineapple", "Strawberry")
```

Fruits array will now be having "Mango", "Apple", "Watermelon", "Pineapple", "Strawberry", "Orange".  

* Remove 3rd index from end of array and add other three fruits

```
Fruits.Splice -3, 1, Array("Watermelon", "Pineapple", "Strawberry")
```
Fruits array will be now having "Mango", "Apple", "Watermelon", "Pineapple", "Strawberry", "Grape", "Orange"