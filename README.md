# excel-column-manager
Save and restore the state of Columns in Tables in Excel.

When dealing with large Tables in Excel with many columns, it would be handy to have a list of presets that show and hide columns depending on your current task. This tool lets you save and restore the visibility and width state of all the columns in a Table.

## 📸 Screenshots
![Screenshot of tool in action](images/Screenshot01.PNG)

## ⭐ Features
- ✅ Save the state (visiblity, width) of Columns in a Table in workbooks persistently (using CustomXML object).
- ✅ Restore saved Column States.
- ✅ Partially restore any Column State to a table if at least one column is present.
- ✅ Re-associate orphaned Column States (i.e., Table name changed).
- ✅ Import and Export Column States manually (using Base64 encoded strings).

## 📝 Notes
- 🚧 [TODO List](Notes/TODO.md)
- States are stored by first Base64 encoding all strings, then creating a comma separated list for properties for each column. Next, we join those individual items into a semi-colon separated list for the entire Table. Lastly, we colon separate this data with the metadata for the entire state, such as the Table Name, the caption, and the last modified date.
- Hidden columns have their ColumnWidth property set to 0. In some cases, it is possible to get the width they will be restored to using the `.Previous.Width` property, but not always.

# 📖 Reference
- [Range.ColumnWidth property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columnwidth)
- [Range.Hidden property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden)

# 🙏 Thanks
- Developed using [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck).