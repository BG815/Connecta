# Connecta

This repository contains the latest version of the Standard Work Form template as well as supporting VBA snippets.

## Standard Work automation macro

The `vba/StandardWorkAutomation.bas` module includes the `UpdateStandardWorkForm` procedure that:

- Reads the Basic (`STDWork_tbl`) and Specific (`SpecificWork_tbl`) tables on the **Standard Work** sheet.
- Groups the tasks by Start of Shift, During Shift, End of Shift, Weekly, and Team Member Specific sections.
- Automatically inserts the required number of rows on the **STDW Form** sheet so the section headers, weekly tasks (starting at `N2`), specific tasks, and notes area stay directly below one another with seven empty rows reserved for notes.

To use the macro:

1. Open `Standard Work Form.xlsm` in Excel.
2. Open the VBA editor (`ALT+F11`), add a new module, and paste the contents of `vba/StandardWorkAutomation.bas` into it.
3. Run `UpdateStandardWorkForm` after editing the Basic or Specific tables to regenerate the form layout.
