# excel-column-manager
Save and restore the state of Columns in Tables in Excel.

When dealing with large tables in Excel with many columns, it would be handy to have a list of presets that show and hide columns depending on your current task. This tool lets you do just that.

## 📸 Screenshots
![Screenshot of tool in action](images/Screenshot01.png)

## ⭐ Features
- ✅ Save the state of a Table's Columns (visiblity, width) persistently in the Workbook (using CustomXML object).
- ✅ Restore saved Column States.
- ✅ Partially restore any Column State to a table if at least one column is present.
- ✅ Re-associate orphaned Column States (i.e., Table name changed).
- ✅ Some useful built-in states such as "Show all (standard)", "Show all (autofit)", and "Show first only".
- ✅ Import and Export Column States manually (using Base64 encoded strings).

## 📝 Notes
- 🚧 [TODO List](Notes/TODO.md)
- ☕ [Dev Notes](Notes/DevNotes.md)

# 🙏 Thanks
- Developed using [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck) and greatly inspired by Mathieu Guindon's blog posts.