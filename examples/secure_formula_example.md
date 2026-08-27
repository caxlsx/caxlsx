## Description

You may use `secure_formulas` to apply Excel's `quotePrefix` style attribute to cells whose string values start with formula-like prefixes (`=`, `+`, `-`, `@`). This prevents Excel from re-evaluating the cell as a formula when a user double-clicks to edit it, without altering the displayed value.

This is a stronger protection than `escape_formulas`, which only prepends an apostrophe at serialization time. With `secure_formulas`, the cell's style in `styles.xml` includes `quotePrefix="1"`, telling Excel to never interpret the content as a formula.

The following are possible:

| Scope     | Example                                                                     | Notes                                                                                      |
|-----------|-----------------------------------------------------------------------------|--------------------------------------------------------------------------------------------|
| Global    | `Axlsx.secure_formulas = true`                                              | Affects workbooks created *after* setting.                                                 |
| Workbook  | `workbook.secure_formulas = true`                                           | Affects child worksheets added *after* setting. Does not affect existing child worksheets. |
| Worksheet | `workbook.add_worksheet(name: 'Name', secure_formulas: true)`               |                                                                                            |
| Worksheet | `worksheet.secure_formulas = true`                                          | Affects child rows/cells added *after* setting. Does not affect existing child rows/cells. |
| Row       | `worksheet.add_row(['+cmd', '-SUM(A1)'], secure_formulas: true)`            | Applies to all cells in the row.                                                           |
| Cell      | `cell.secure_formulas = true`                                               |                                                                                            |

Non-string types (numbers, dates, booleans) are never affected.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(name: 'Secure Formulas', secure_formulas: true) do |sheet|
  # These cells get quotePrefix applied (formula-like prefixes)
  sheet.add_row ["+cmd|'/C calc'!A0", "-SUM(A1)", "@SUM(A1)", "=1+1"]

  # Non-string and non-prefixed cells are unaffected
  sheet.add_row [42, "safe text", 3.14]

  # Per-cell override
  sheet.add_row ["=HYPERLINK(...)"], secure_formulas: false
end

p.serialize 'secure_formula_example.xlsx'
```

## Output

Cells with formula-like prefixes will display their literal text and cannot be re-evaluated as formulas, even when edited by the user.
