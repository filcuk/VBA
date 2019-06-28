# Personal VBA KB

*I'm still learning Markup, don't think about it too much.*

### Determine uninitialised array
`If Not Not myArray Then Debug.Print UBound(myArray) Else Debug.Print "Not initialised"`  
*This quirk requires double negative to work correctly.*

### DGET()

### Selection
`ActiveWindow.RangeSelection`
- How is this different from ActiveSheet.Selection?
> Glad you asked. This returns selected range **even** if an object (such as graphic) is selected.

### Division
`3 / 2`  
> 1.5  
`3 \ 2`  
> 1  

### UDF param tooltip
- You can press `CTRL+SHIFT+A` to show all parameter vars
- You can find and edit properties of your UDF in the VBA IDE object browser

### Variable Types
```
Dim iNumber%     'Integer  
Dim lAverage&    'Long  
Dim sngTotal!    'Single  
Dim dbTotal#     'Double  
Dim cProfit@     'Currency  
Dim sFirstName$  'String  
Dim llDiscount^  'LongLong on 64 bit
```

### .SpecialCells()

First of all, be aware that **if the input is range of 1 cell, the function considers UsedRange instead!**
This does not apply to types 5, 6, 9 or 10 (more on that below).

Argument 1: Type (Required)

*Note that 7, 8, 13 or 14 seem to ignore the value argument.*
*Below is the full list of arguments, not all officially documented.*

| Type number       | Equivalent                                    | xlCellType Enum |
|-------------------|-----------------------------------------------|-----------------|
| .SpecialCells(1)  | .SpecialCells(xlCellTypeComments)             | -4144           |
| .SpecialCells(2)  | .SpecialCells(xlCellTypeConstants)            | 2               |
| .SpecialCells(3)  | .SpecialCells(xlCellTypeFormulas)             | -4123           |
| .SpecialCells(4)  | .SpecialCells(xlCellTypeBlanks)               | 4               |
| .SpecialCells(5)  | .CurrentRegion                                |                 |
| .SpecialCells(6)  | .CurrentArray                                 |                 |
| .SpecialCells(7)  | .RowDifferences                               |                 |
| .SpecialCells(8)  | .ColumnDifferences                            |                 |
| .SpecialCells(9)  | .Precedents                                   |                 |
| .SpecialCells(10) | .Dependents                                   |                 |
| .SpecialCells(11) | .SpecialCells(xlCellTypeLastCell)             | 11              |
| .SpecialCells(12) | .SpecialCells(xlCellTypeVisible)              | 12              |
| .SpecialCells(13) | .SpecialCells(xlCellTypeAllFormatConditions)  | -4172           |
| .SpecialCells(14) | .SpecialCells(xlCellTypeAllValidation)        | -4174           |
|                   | .SpecialCells(xlCellTypeSameFormatConditions) | -4173           |
|                   | .SpecialCells(xlCellTypeSameValidation)       | -4175           |

Argument 2: Value (Optional)

*Note that this value is additive, so you can use any number of values at once.*

| Enum         | Literal |
|--------------|---------|
|  xlNumbers   | 1       |
| xlTextValues | 2       |
| xlLogical    | 4       |
| xlErrors     | 16      |
