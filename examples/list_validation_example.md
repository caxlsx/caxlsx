## Description

You could validate data based on an explicit list or a given range of cells.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(name: 'Basic Worksheet') do |sheet|
  sheet.add_row ['First', 'Second', 'Third']
  sheet.add_row [1, 6, 11]

  sheet.add_data_validation('A2:A2',
    type: :list,
    formula1: 'A1:C1',
    showDropDown: false,
    showErrorMessage: true,
    errorTitle: '',
    error: 'Only values from A1:C1',
    errorStyle: :stop,
    showInputMessage: true,
    promptTitle: '',
    prompt: 'Only values from A1:C1')

  sheet.add_data_validation('B2:B2',
    type: :list,
    formula1: '"Red, Orange, NavyBlue"',
    showDropDown: false,
    showErrorMessage: true,
    errorTitle: '',
    error: 'Please use the dropdown selector to choose the value',
    errorStyle: :stop,
    showInputMessage: true,
    prompt: 'Choose the value from the dropdown')
end

p.serialize 'list_validation_example.xlsx'
```

## Output

Validating by cells:

![Output](images/list_validation_example_1.png "Output")

Validating by list:

![Output](images/list_validation_example_2.png "Output")
