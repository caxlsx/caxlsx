## Description

Book views let you specify which sheet the show as active when the user opens the work book as well as a bunch of other.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(name: 'Basic Worksheet') do |sheet|
  sheet.add_row ['First', 'Second', 'Third']
  sheet.add_row [1, 2, 3]
end

wb.add_worksheet(name: 'Second Worksheet') do |sheet|
  sheet.add_row ['First', 'Second', 'Third']
  sheet.add_row [1, 2, 3]
end

# The horizontal scrollbar will be smaller
# The second tab will be selected
wb.add_view tab_ratio: 800, active_tab: 1

p.serialize 'active_tab_example.xlsx'
```

## Output

![Output](images/active_tab_example.png "Output")
