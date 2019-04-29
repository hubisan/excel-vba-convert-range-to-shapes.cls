# Convert an Excel range to shapes

Convert texts, borders and fills of an Excel range to shapes. The shapes version tries to be as identical to the Excel range as possible.

Example:



Use cases:
- Easily copy an Excel range as it is to PowerPoint.
- Create dashboards in Excel and having tables below or next to each another with different column widths and row heights.

_Remark for texts:_

_Make sure to explicitly set the text alignment to be able to copy the alignment for the shapes. Else it can lead to not identically aligned texts. If a text spans multiple cells the cells need to be merged to make the shape cover the same range. Else the text might get trunctuated._

## Installation

- *Download the workbook*: You can download [ConvertRangeToShapes.xlsm](../../raw/master/ConvertRangeToShapes.xlsm). This workbook inlcudes the class and a module with two examples on how to use the class.
- *Or copy the code*: Copy the class code starting at line 10 `#If VBA7` from [raw](../../raw/master/ConvertRangeToShapes.cls) and paste into a new class in your workbook.
- *Or get the repository*: Clone or [download](../../archive/master.zip) the repository and import the file stopwatch.cls into your workbook to use the class.

### Usage

Methods and properties (see docstrings in VBA code for more details):

- `.start`: start or restart the timer
- `.Elapsed_ms`: get the elapsed time in milliseconds (like 1004) As Double
- `.Elapsed_sec(Optional number_of_digits_after_decimal As Integer = 3)`: get the elapsed time in seconds as Double.  
  You can round to a number of digits after the decimal by providing the optional parameter, the default is 3 = millliseconds.
- `.stop_it`: stops the timer, rarely using this myself.

Example usage:

```vb
Sub convertAll()
    Dim x As New ConvertRangeToShapes
    ' Convert borders, fill and text. Ignore white fill.
    x.convertAll ThisWorkbook.Worksheets("-").Range("B8:C12"), "all", RGB(255, 255, 255)
End Sub

Sub convertBordersOnly()
    Dim x As New ConvertRangeToShapes
    
    ' Convert borders only. Delete and select is only run automatically when using convertAll
    x.deleteShapes ThisWorkbook.Worksheets("-"), "exampleBorders"
    x.convertBorders ThisWorkbook.Worksheets("-").Range("B8:C12"), "exampleBorders"
    x.selectShapes ThisWorkbook.Worksheets("-"), "exampleBorders"
End Sub

Sub convertFillsOnly()
    Dim x As New ConvertRangeToShapes
    
    ' Convert fills only.
    x.deleteShapes ThisWorkbook.Worksheets("-"), "exampleFills"
    x.convertFills ThisWorkbook.Worksheets("-").Range("B8:C12"), , "exampleFills"
    x.selectShapes ThisWorkbook.Worksheets("-"), "exampleFills"
End Sub

Sub convertFillsOnlyIgnoreWhite()
    Dim x As New ConvertRangeToShapes
    
    ' Convert fills only. Use setName to not have to repeat the name.
    x.setShapesName "exampleFills1"
    x.deleteShapes ThisWorkbook.Worksheets("-")
    x.convertFills ThisWorkbook.Worksheets("-").Range("B8:C12"), RGB(255, 255, 255)
    x.selectShapes ThisWorkbook.Worksheets("-")
End Sub

Sub convertTextsOnly()
    Dim x As New ConvertRangeToShapes
    
    ' Convert texts only.
    x.deleteShapes ThisWorkbook.Worksheets("-"), "exampleTexts"
    x.convertTexts ThisWorkbook.Worksheets("-").Range("B8:C12"), "exampleTexts"
    x.selectShapes ThisWorkbook.Worksheets("-"), "exampleTexts"
End Sub


Sub deleteAll()
    Dim x As New ConvertRangeToShapes
    
    ' Using wildcard to delete all.
    x.deleteShapes ThisWorkbook.Worksheets("-"), "example*"
End Sub
```

## Remark for the workbook

The class module gets automatically exported when saving the Excel file. To stop this behavior change the constant in the module "autoopen" to "False":
```vb
Private Const is_in_development As Boolean = False
```

## Contributing

If you find a bug, please create a new issue. Pull requests are also welcome.

## Contributors

- [Daniel Hubmann](https://github.com/hubisan) (Author)

## License

Copyright (c) 2019 Daniel Hubmann. Licensed under [MIT](LICENSE).
