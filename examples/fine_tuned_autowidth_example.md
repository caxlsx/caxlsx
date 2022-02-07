## Description

This is an example of how to use font_scale_divisor and bold_font_multiplier to fine-tune the automatic column width calculation.

## Code

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook
wb.font_scale_divisor = 11.5
wb.bold_font_multiplier = 1.05
s = wb.styles
sm_bold = s.add_style b: true, sz: 5
m_bold = s.add_style b: true, sz: 11
b_bold = s.add_style b: true, sz: 16
ub_bold = s.add_style b: true, sz: 21
sm_text = s.add_style sz: 5
m_text = s.add_style sz: 11
b_text = s.add_style sz: 16
ub_text = s.add_style sz: 21

wb.add_worksheet(name: 'Line Chart') do |sheet|
  sheet.add_row(['Lorem', 'Lorem', 'Lorem', 'Lorem', 'Lorem', 'Lorem', 'Lorem', 'Lorem', 'WWW', 'www', 'iii', 'iii',
                 'Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed',
                 'Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed'],
                style: [sm_bold, m_bold, b_bold, ub_bold, sm_text, m_text, b_text, ub_text,
                        b_bold, b_text, b_bold, b_text, b_bold, b_text])
end

p.serialize 'fine_tuned_autowidth_example.xlsx'
```

## Output

![Output](images/fine_tuned_autowidth_example.png "Output")