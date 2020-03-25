## Description

There is a way to add styles to a whole column.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

s = wb.styles
percent = s.add_style num_fmt: 9, bg_color: '0000FF'

wb.add_worksheet(name: 'Custom column types') do |sheet|
  sheet.add_row ['A', 'B', 'Percent', 'Hidden', 'E']
  sheet.add_row [1, 2, 0.3, 4, 5.0]
  sheet.add_row [1, 2, 0.2, 4, 5.0]
  sheet.add_row [1, 2, 0.1, 4, 5.0]

  # Apply the percent style to the column at index 2 skipping the first row.
  sheet.col_style 2, percent, row_offset: 1
  sheet.column_info[3].hidden = true
end

p.serialize 'column_styles_example.xlsx'
```

## Output

![Output](images/column_styles_example.png "Output")
