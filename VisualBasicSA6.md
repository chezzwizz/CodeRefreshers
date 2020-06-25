# Keywords

`Dim`: The `Dim` keyword is used to explicitly declare a variable. I have read that it means dimension, though I like to rember it as an acronym for "Declare in Memory" since this is essentially what happens when you run your program and a `Dim` statement is encountered.

# Options
_Applies to: VBA_

The Visual Basic language has some specifiers which give you control over how the compiler interprets your code. These are similar to preprocessor directives, but are different in how they are declared. Where a VB preprocessor directive would use the # as a prefix, the Option statement is simply `Option opt [arg]`, where `arg` is only available for some commands. The most common options follow:

* `Explicit`: Declaring `Option Explicit` forced the compiler to check for explicit variable decleration using the `Dim` (for dimension) keyword. Variables implicitly declared are flagged with an error and prevent the program from being compiled. Enabling this option works great for making sure you are spelling your variable names every time you use them.
* `Compare`: This is an option that specifies the default string comparison method. The `Option Compare` decleration also requires an argument:
  |String comparison option|Description|
  |:---|:---|
  |Database| |
  |Binary| |
  |Text| |
* `Base`: Declaring `Option Base 1` defaults all arrays __within the module or class file__ to have a lower boundry of 1 instead of 0. One thing to be aware of is that some of the built in objects have been constructed using a lower boundry of 0 and will still only work as expected when used that way. If you plan for maximum compatibility with all built in procedures, it is advisable to work with base 0. While `Option Base 1` can help readability, it can make it slightly confusing when you start importing external modules.

# Arrays
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

# Sources

_Links connect to [GoodReads.com](https://www.goodreads.com) search page_

1. [Mastering VBA for Microsoft Office 365](https://www.goodreads.com/search?utf8=%E2%9C%93&query=978-1-119-57933-5) by [Richard Mansfield](https://www.goodreads.com/author/show/18687941.Richard_Mansfield?from_search=true&from_srp=true)
2. [Access 2019 Bible](https://www.goodreads.com/search?utf8=%E2%9C%93&query=978-1-119-51475-6) by [Michael Alexander](https://www.goodreads.com/author/show/6526935.Michael_Alexander?from_search=true&from_srp=true) and [Richard Kusleika](https://www.goodreads.com/author/show/14301914.Richard_Kusleika?from_search=true&from_srp=true)
