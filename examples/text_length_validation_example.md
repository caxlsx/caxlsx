## Description

You could make validations based on text length

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(name: 'Basic Worksheet') do |sheet|
  sheet.add_row ['First', 'Second', 'Third']
  sheet.add_row ['Some', 'text', 'to validate']

  sheet.add_data_validation('A2:C2',
    type: :textLength,
    operator: :lessThan,
    formula1: '10',
    showErrorMessage: true,
    errorTitle: 'Text is too long',
    error: 'Max text length is 10 characters',
    errorStyle: :stop,
    showInputMessage: true,
    promptTitle: 'Text length',
    prompt: 'Max text length is 10 characters')
end

p.serialize 'text_length_validation_example.xlsx'
```

## Output

![Output](images/text_length_validation_example.png "Output")
