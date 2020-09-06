## Description

Conditional format example: Between. Use array for formula.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

s = wb.styles
high = s.add_style bg_color: 'FF428751', type: :dxf

wb.add_worksheet(name: 'Conditional') do |sheet|
  # Use 10 random number
  sheet.add_row Array.new(10) { (rand * 10).floor }

  sheet.add_conditional_formatting('A1:J1',
    type: :cellIs,
    operator: :between,
    formula: ['3', '7'],
    dxfId: high,
    priority: 1)
end

p.serialize 'conditional_formatting_between_example.xlsx'
```

## Output

![Output](images/conditional_formatting_between_example.png "Output")
