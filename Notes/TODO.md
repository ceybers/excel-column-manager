# TODO
## Target ListObject
- [x] Show name of Target ListObject.
- [ ] Give states names, defaulting to e.g., "(x/y) Untitled state (2023/01/01 09:00)".
- [x] Allow saving current State, if not already saved.
- [ ] Check object references work OK with multiple workbooks.
## TreeView
- [x] List all saved Column States.
- [x] List current (unsaved) State.
- [x] Auto-selects a saved state if it matches the current state.
- [x] Target ListObject is top of the tree.
- [x] Orphaned items are bottom of the tree.
- [ ] Consider adding meta states for "Show all (default)", "Show all (autofit)", and "Show first only".
## ListView
- [x] Show list of columns in a stored State.
- [x] Indicate Hidden property with icon.
- [x] Indicate Width property.
- [x] Indicate whether column exists in target ListObject.
## Options
- [x] Option to Associate on applying orphan.
- [x] Option to Close on Apply.
- [x] Option to Allow/Prohibit displaying Partial matches.
- [x] Option to Allow/Prohibit applying Partial matches.
- [x] Option to hide columns in target that were not in the applied state (i.e., hide unless specifically stated to Show). (Default behaviour is to leave columns not in the state as-is).
- [x] Filter to only show existing columns in State preview.
- [x] Default values if no CustomXML found.
## Features
- [x] Apply selected state.
- [x] Import/export manually using Base64 string.
- [x] Prune orphaned states.
- [x] Remove selected state.
- [x] Remove all states.
- [ ] Serialize Freeze state of worksheet.
- [ ] Consider serializing Outline levels.
- [ ] Store generic column states in User-level persistence (i.e., for applying to new/unsaved files).
- [ ] Dedicated Memory Recall/Store state for direct use from Ribbon (i.e., no UserForm).