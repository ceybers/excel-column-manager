# TODO
## Target (Top)
- [x] Show Target ListObject name.
- [x] Allow saving current State, if not already saved.
## States (LHS)
- [x] List all saved Column States.
- [ ] List current (unsaved) State.
- [ ] Auto-selects a saved state if it matches the current state.
- [x] Target ListObject is top of the tree.
- [x] Orphaned items are bottom of the tree.
- [ ] Consider adding meta states for "Show all (default)", "Show all (autofit)", and "Show first only".
## Selected (RHS)
- [x] Show list of columns in a stored State.
- [x] Indicate Hidden property with icon.
- [x] Indicate Width property.
- [x] Indicate whether column exists in target ListObject.
## Features
- [ ] Give states names, defaulting to e.g., "(x/y) Untitled state (2023/01/01 09:00)".
- [x] Apply selected state.
- [ ] Import/export manually using Base64 string.
- [ ] Prune orphaned states.
- [x] Remove selected state.
- [x] Remove all states.
## Options (sub UserForm)
- [ ] Create child ViewModel to contain all these flags.
- [ ] Option to Associate on applying orphan.
- [ ] Option to Close on Apply.
- [ ] Option to Allow/Prohibit displaying Partial matches.
- [ ] Option to Allow/Prohibit applying Partial matches.
- [ ] Option to hide columns in target that were not in the applied state 
  - (i.e., hide unless specifically stated to Show).
  - Default behaviour is to leave columns not in the state as-is.
- [ ] Filter to only show existing columns in State preview.
- [ ] Search Match Case
- [ ] Search Match Whole Word
- [ ] Default values for all options if no CustomXML found.
## Nice-to-Haves
- [ ] Use nicer icons. Probably use .PNG and store them in a hidden UserForm.
- [ ] Serialize Freeze Pane state of worksheet.
- [ ] Consider serializing Outline levels.
- [ ] Store generic column states in User-level persistence (i.e., for applying to new/unsaved file)
- [ ] Status bar with history log.
- [ ] Implement Table Picker to swap Target while UserForm is active.
## Admin
- Check object references work OK with multiple workbooks.
- Implement unit tests for everything.
- Document code.
- Update README/Git page. 