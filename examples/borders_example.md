## Description

Axlsx defines a thin border style, but you can easily create and use your own.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

s = wb.styles
red_border = s.add_style border: { style: :thick, color: 'FFFF0000', edges: [:left, :right] }
blue_border = s.add_style border: { style: :thick, color: 'FF0000FF' }

complex_border = workbook.styles.add_style(
  border: [
    { style: :thin, color: '000' },
    { style: :double, edges: [:top, :bottom] },
    { style: :thick, edges: [:left, :right] },
  ]
)

wb.add_worksheet(name: 'Custom Borders') do |sheet|
  sheet.add_row ['wrap', 'me', 'up in red'], style: red_border
  sheet.add_row [1, 2, 3], style: blue_border
  sheet.add_row [4, 5, 6], style: complex_border
end

p.serialize 'borders_example.xlsx'
```

## Output

![Output](images/borders_example.png "Output")
