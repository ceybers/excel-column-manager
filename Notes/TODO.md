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
- [ ] Give states names, defaulting to e.g., "(x/y) Untitled state (2023/01/01 09:00)".
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
- [x] Search Match Case
- [x] Search Match Whole Word
- [x] Default values for all options if no CustomXML found.
## Nice-to-Haves
- [x] Serialize Freeze Pane state of worksheet.
- [x] Consider serializing Outline levels.
- [ ] Inform user if search returned no results instead of just showing an empty tree/list.
- [ ] Use nicer icons. Probably use .PNG and store them in a hidden UserForm.
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