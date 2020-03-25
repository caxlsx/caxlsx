## Description

You could serialize the package into a stream instead of writing to a file directly.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(name: 'Basic Worksheet') do |sheet|
  sheet.add_row ['First', 'Second', 'Third']
  sheet.add_row [1, 2, 3]
end

stream = p.to_stream
File.open('stream_example.xlsx', 'w') { |f| f.write(stream.read) }
```

## Output

No difference should be visible
