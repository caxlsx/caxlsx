## Description

Conditional format example: Text equal.

1. You must specify `:containsText` for both type and operator
2. You must craft a formula to match what you are looking for. In this case we only need to reference the top-left cell of the range.
3. Formula can be very vender specific. You will want to test extensively if interoperability beyond excel is a concern.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

s = wb.styles
profit = s.add_style bg_color: 'FF428751', type: :dxf

wb.add_worksheet(name: 'Text Matching Conditional') do |sheet|
  # Use 10 random number
  sheet.add_row Array.new(10) { ['Profit', 'Loss'].sample }

  sheet.add_conditional_formatting('A1:J1',
    type: :containsText,
    operator: :containsText,
    formula: 'NOT(ISERROR(SEARCH("Profit",A1)))',
    dxfId: profit,
    priority: 1)
end

p.serialize 'conditional_formatting_text_equal_example.xlsx'
```

## Output

![Output](images/conditional_formatting_text_equal_example.png "Output")
