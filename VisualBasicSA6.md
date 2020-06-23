# Basic Arrays
_Applies to: VB6, VBA_

### Declaring an array

```
  ' Declares a dynamic array of type Variant
  Dim A()
  
  ' Declares a typed array with a single dimension defined with 11 elements
  ' set to the integer default of 0. By default, this array has a lower
  ' bound (first index) of 0 and an upper bound of 10 unless 'Option Base 1'
  ' is declared, in which case there are only 10 elements (1 to 10). In any
  ' case, the upper bound is always set to what is specified between the parenthesis.
  Dim A(10) as Integer
```
___NOTE:___ As a side note, many popular C inspired programming languages use 0 as the first index. One explination for this is the idea that as a place holder, zero is a perfectly good symbol while in a mathematical sense, it means a lack of value. IMHO, it is a good habit to get into of being aware of what your context is. One way to practice this is to use VB:)

```
  ' Declares an array of type Variant with a single dimension and 10
  ' elements with a lower bound of 1 regardless of if 'Option Base 1'
  ' is declared or not.
  Dim A(1 to 10)
```
