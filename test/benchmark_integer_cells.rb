# frozen_string_literal: true

# Run with: ROWS=10000 bundle exec ruby test/benchmark_integer_cells.rb
$LOAD_PATH.unshift File.expand_path('../lib', __dir__)
require 'axlsx'
require 'digest'
require 'json'

def measure_cells(rows, values)
  GC.start
  allocated = GC.stat(:total_allocated_objects)
  wall = Process.clock_gettime(Process::CLOCK_MONOTONIC)
  cpu = Process.clock_gettime(Process::CLOCK_PROCESS_CPUTIME_ID)
  package = Axlsx::Package.new
  package.use_autowidth = false
  sheet = package.workbook.add_worksheet(name: 'Numbers')
  rows.times { sheet.add_row(values) }
  xml = sheet.to_xml_string
  {
    rows: rows, cells: rows * values.length,
    wall: Process.clock_gettime(Process::CLOCK_MONOTONIC) - wall,
    cpu: Process.clock_gettime(Process::CLOCK_PROCESS_CPUTIME_ID) - cpu,
    allocations: GC.stat(:total_allocated_objects) - allocated,
    xml_sha256: Digest::SHA256.hexdigest(xml), types: sheet.rows.first.cells.map(&:type)
  }
end

rows = Integer(ENV.fetch('ROWS', '3000'))
inputs = {
  integers: [0, -37, 123, 2**80, -(2**80)] * 4,
  mixed: [0, -37, 123, 2**80, 0.125, -17.75, '0012', '1e308', '1e-308', Float::MIN, Float::INFINITY, nil, true, false, 'text', '2008-08-30T01:45:36.123']
}

inputs.each do |label, values|
  3.times do |run|
    puts JSON.generate(case: label, run: run + 1, **measure_cells(rows, values))
  end
end
