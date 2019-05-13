# Personal VBA KB

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
