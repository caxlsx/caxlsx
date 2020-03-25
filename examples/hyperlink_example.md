## Description

You could create an internal or external hyperlink. The ref could be a string referencing to an existing cell or a cell object.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(name: 'Hyperlinks') do |sheet|
  # external references
  sheet.add_row ['axlsx']
  sheet.add_hyperlink location: 'https://github.com/caxlsx/caxlsx', ref: 'A1'
  # internal references
  sheet.add_hyperlink location: "'Next Sheet'!A1", ref: 'A2', target: :sheet
  sheet.add_row ['next sheet']
end

wb.add_worksheet(name: 'Next Sheet') do |sheet|
  sheet.add_row ['hello!']
end

p.serialize 'hyperlink_example.xlsx'
```

## Output

![Output](images/hyperlink_example.png "Output")
