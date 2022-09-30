## Description

Shows how to append styles after rows have been created using `worksheet.add_style`

## Code

```ruby
p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet do |sheet|
  sheet.add_row
  sheet.add_row ["", "Product", "Category",  "Price"]
  sheet.add_row ["", "Butter", "Dairy",      4.99]
  sheet.add_row ["", "Bread", "Baked Goods", 3.45]
  sheet.add_row ["", "Broccoli", "Produce",  2.99]

  sheet.add_style "B2:D2", b: true
  sheet.add_style "B2:B5", b: true
  sheet.add_style "B2:D2", bg_color: "95AFBA"
  sheet.add_style "B3:D5", bg_color: "E2F89C"
  sheet.add_style "D3:D5", alignment: { horizontal: :left }
  sheet.add_style ["C3:C4", "D3:D4"], fg_color: "00FF00"
nd

p.serialize "append_styles.xlsx"
```

## Output

![Output](images/append_styles.png "Output")
