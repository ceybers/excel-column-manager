# TODO
## Target (Top)
- [x] Show Target ListObject name.
- [x] Allow saving current State, if not already saved.
## States (LHS)
- [x] List all saved Column States.
- [x] List current (unsaved) State.
- [x] Auto-selects a saved state if it matches the current state.
- [x] Target ListObject is top of the tree.
- [x] Orphaned items are bottom of the tree.
- [x] Entering search Text Box should remove watermark text.
- [x] Consider adding meta states for "Show all (default)", "Show all (autofit)", and "Show first only".
## Selected (RHS)
- [x] Show list of columns in a stored State.
- [x] Indicate Hidden property with icon.
- [x] Indicate Width property.
- [x] Indicate whether column exists in target ListObject.
## Features
- [x] Give states names, defaulting to e.g., "(x/y) Untitled state (2023/01/01 09:00)".
- [x] Apply selected state.
- [x] Import manually using Base64 string.
- [x] Export manually using Base64 string.
- [x] Prune orphaned states.
- [x] Remove selected state.
- [x] Remove all states.
- [x] Serialize Width of hidden columns.
## Bugs
- [x] Closing Options UserForm closes main Explorer UserForm.
  - Main UserForm was accidentally set to non-modal.
- [x] RemoveOrphan fails - check if we're trying to cast a Builtin State to a normal state.
- [x] "No search results" is displayed when there are no saved states, even if the user is not searching.
- [x] Production build includes testing states even in a blank new Workbook.
- [x] Default caption (comma separated list of Column names) on Tables with many columns is unusuable.
  - Only list visible columns, trim ellipsis each column name, limit to first `n` columns.
- [x] Increase width of ListView column for the index of a column so it can contain double digits.
- [x] RemoveAll removes Built-in States.
- [x] Check if Worksheet is Protected before trying to make changes.
- [x] Pressing Esc from the Rename State dialog resets caption to `GetAutoCaption`.
- [x] Pressing Esc in "Import Serial State" dialog returns error (malformed serial).
- [x] Auto-fit ListView column width (i.e., Parent.Width - 4) still causes a Horizontal scrollbar if there is a Vertical Scrollbar active.
- [x] Freeze panes in Built-in states assume ListObject starts at cell A1.
## Options (sub UserForm)
- [x] Create child ViewModel to contain all these flags.
- [x] Show/hide orphaned states.
- [x] Option to Associate on applying orphan.
- [x] Option to Close on Apply.
- [x] Option to Allow/Prohibit displaying Partial matches.
- [x] Option to Allow/Prohibit applying Partial matches.
- [x] Option to hide columns in target that were not in the applied state 
  - (i.e., hide unless specifically stated to Show).
  - Default behaviour is to leave columns not in the state as-is.
- [x] Show/hide non-existing columns in State preview.
- [ ] Show/hide matching but hidden columns in States.
- [x] Search Match Case
- [x] Search Match Whole Word
- [x] Default values for all options if no CustomXML found.
## Nice-to-Haves
- [x] Serialize Freeze Pane state of worksheet.
- [x] Consider serializing Outline levels.
  - [ ] Check if it is necessary to set all Outline Levels to 1 before applying them from a state, i.e. to prevent a malformed Outline state midway through Applying.
- [x] Inform user if search returned no results instead of just showing an empty tree/list.
- [x] Cancel on pressing Escape key (hidden Cancel=true button). 
- [x] Use nicer icons. Probably use .PNG and store them in a hidden UserForm.
- [x] Use `ScreenUpdating` when Applying a state for speed.
- [x] Use `ScreenUpdating` when creating ViewModel (initial show of UI) for speed.
- [x] Double-clicking a Built-in State will Apply it (since we can't rename them).
- [ ] Store generic (non-Workbook specific) column states in User-level persistence
  - (i.e., for applying to new/unsaved file)
- [ ] Status bar with history log.
- [ ] Implement Table Picker to swap Target while UserForm is active.
- [ ] Clean up `StateCast` workarounds to handle both `ColumnState` and built-in `IState` in the same Collection.

## Admin
- Check object references work OK with multiple workbooks.
- Implement unit tests for everything.
- Document code.
- Update README/Git page. 