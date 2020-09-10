## Description

You could set the font


## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

s = wb.styles
courier = s.add_style font_name: 'Courier'

wb.add_worksheet(name: 'Custom Styles') do |sheet|
  sheet.add_row ['First', 'Second', 'Third'], style: [nil, courier, nil]
end

p.serialize 'font_example.xlsx'
```

## Output

![Output](images/font_example.png "Output")
