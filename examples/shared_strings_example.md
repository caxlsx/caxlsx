## Description

To work with Apple Numbers, you must set `use_shared_strings` to true. This is most conveniently done just before rendering by calling `p.use_shared_strings = true` prior to serialization.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(name: 'Basic Worksheet') do |sheet|
  sheet.add_row ['First', 'Second', 'Third']
  sheet.add_row [1, 2, 3]
end

p.use_shared_strings = true
p.serialize 'shared_strings_example.xlsx'
```

## Output

No difference should be visible
