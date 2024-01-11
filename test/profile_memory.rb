#!/usr/bin/env ruby -s
# frozen_string_literal: true

$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'memory_profiler'

# Axlsx.trust_input = true

row = []
input1 = (32..126).to_a.pack('U*').chars.to_a # these will need to be escaped
input2 = (65..122).to_a.pack('U*').chars.to_a # these do not need to be escaped
10.times { row << input1.join }
10.times { row << input2.join }

report = MemoryProfiler.report do
  p = Axlsx::Package.new
  p.workbook.add_worksheet do |sheet|
    10_000.times do
      sheet << row
    end
  end
  p.serialize("example_memory.xlsx", zip_command: 'zip')
end
report.pretty_print

File.delete("example_memory.xlsx")
