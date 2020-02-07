require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

#   Each style only needs to be declared once in the workbook.
s = wb.styles
title_cell = s.add_style sz: 14, alignment: { horizontal: :center }
black_cell = s.add_style sz: 10, alignment: { horizontal: :center }
blue_link = s.add_style fg_color: '0000FF'

# Create summary sheet
sum_sheet = wb.add_worksheet name: 'Summary'
sum_sheet.add_row ['Test Results'], sz: 16
sum_sheet.add_row
sum_sheet.add_row
sum_sheet.add_row ['Note: Blue cells below are links to the Log sheet'], sz: 10
sum_sheet.add_row
sum_sheet.add_row ['ID', 'Type', 'Match', 'Mismatch', 'Diffs', 'Errors', 'Result'], style: title_cell

sum_sheet.column_widths 10, 10, 1 0, 11, 10, 10, 10

# Starting data row in summary spreadsheet (after header info)
current_row = 7

# Create Log Sheet
log_sheet = wb.add_worksheet name: 'Log'
log_sheet.column_widths 10
log_sheet.add_row ['Log Detail'], sz: 16
log_sheet.add_row

# Add rows to summary sheet
(1..10).each do |sid|
  sum_sheet.add_row [sid, 'test', '1', '2', '3', '4', '5'], style: black_cell

  if sid.odd?
    link = "A#{current_row}"
    sum_sheet[link].style = blue_link
    sum_sheet.add_hyperlink location: "'Log'!A#{current_row}", target: :sheet, ref: link
  end

  current_row += 1
end

p.serialize 'where_is_my_color.xlsx'
