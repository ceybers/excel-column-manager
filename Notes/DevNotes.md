# Development Notes for Excel Column Manager Tool
## Notes
- Should replace the use of `Collection` objects with a more suitable collection class.
  - Lots of poor code due to having to check if items Exist and adding/removing and iterating/mapping.
- MVVM implementation should be using `PropertyChanged` events instead of the View polling the ViewModel on every action.
- Likewise, should be using the `Action` pattern for `Can`/`Do`.
- Simple boolean Options could be separated into their own ViewModel.
## Serializing states
- States are stored by first Base64 encoding all strings, then creating a comma separated list for properties for each column. Next, we join those individual items into a semi-colon separated list for the entire Table. Lastly, we colon separate this data with the metadata for the entire state, such as the Table Name, the caption, and the last modified date.
  - This is to avoid having to deal with escaping characters and to easily use the `Split()` function.
  - e.g. (without Base64 encoding):
> `Table1:State Caption:Column1,Column1Width;Column2,Column2Width;Column3,Column3Width`
- For Tuple properties, e.g. Row and Column, I will propably separate them using a period, e.g., `Row 1 Column 3` becomes `1.3`.
## Hidden Columns and Column Width
- Hidden columns have their ColumnWidth property set to 0. 
- ~~In some cases, it is possible to get the width they will be restored to using the `.Previous.Width` property, but not always.~~
- Correct way is to disable `ScreenUpdating`, unhide the column, copy the `ColumnWidth`, then hide it again. Should occur fast enough that the user won't notice.
## Serializing Freeze Panes
- Properties `ActiveWindow.FreezePanes`, `ActiveWindow.SplitRow`, and `ActiveWindow.SplitColumn`.
- `ActiveWindow.Split` exists as well.
## Serializing Outlines
- `Activesheet.Outline.SummaryColumn = xlSummaryOnRight`
- `Activesheet.Outline.SummaryRow = xlSummaryBelow`
- `Selection.EntireColumn.OutlineLevel`
- Default OutlineLevel is 1, not 0. This can be set while selection is inside a ListObject.
- For setting ActiveSheet Summary Row/Col, remember to move the selection outside of any ListObject or it will throw an error.
>```vb
>    Dim UsedRange As Range
>    Set UsedRange = Target.UsedRange
>    Target.Cells(1, UsedRange.Columns.Count + 1).Select
>```
## 📖API References
- [Range.ColumnWidth property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columnwidth)
- [Range.Hidden property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden)
- [Range.Previous property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.range.previous)
- [Window.FreezePanes property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.window.freezepanes)