#!/usr/bin/env ruby -s

$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'csv'
require 'benchmark'
Axlsx::trust_input = true
row = []
input = (32..126).to_a.pack('U*').chars.to_a
20.times { row << input.shuffle.join }
times = 3000
Benchmark.bmbm(30) do |x|
  x.report('axlsx_noautowidth') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.use_autowidth = false
    p.serialize("example_noautowidth.xlsx")
  end

  x.report('axlsx') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.serialize("example_autowidth.xlsx")
  end

  x.report('axlsx_shared') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.use_shared_strings = true
    p.serialize("example_shared.xlsx")
  end

  x.report('axlsx_stream') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    s = p.to_stream()
    File.open('example_streamed.xlsx', 'wb') { |f| f.write(s.read) }
  end

  x.report('csv') do
    CSV.open("example.csv", "wb") do |csv|
      times.times do
        csv << row
      end
    end
  end
end
File.delete("example.csv", "example_streamed.xlsx", "example_shared.xlsx", "example_autowidth.xlsx", "example_noautowidth.xlsx")
