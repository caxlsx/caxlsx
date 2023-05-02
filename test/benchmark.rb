#!/usr/bin/env ruby -s
# frozen_string_literal: true
$:.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'csv'
require 'benchmark'
# Axlsx.trust_input = true
Axlsx.skip_validations = true
row = []
input1 = (32..126).to_a.pack('U*').chars.to_a # these will need to be escaped
input2 = (65..122).to_a.pack('U*').chars.to_a # these do not need to be escaped
10.times { row << input1.shuffle.join }
10.times { row << input2.shuffle.join }
times = 3000

Benchmark.bmbm(30) do |x|
  x.report('axlsx_noautowidth') {
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
  }

  x.report('axlsx_autowidth') {
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.serialize("example_autowidth.xlsx")
  }

  x.report('axlsx_shared') {
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
  }

  x.report('axlsx_stream') {
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
  }

  x.report('axlsx_zip_command') {
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.serialize("example_zip_command.xlsx", zip_command: 'zip')
  }

  x.report('csv') {
    CSV.open("example.csv", "wb") do |csv|
      times.times do
        csv << row
      end
    end
  }
end
File.delete("example.csv", "example_streamed.xlsx", "example_shared.xlsx", "example_autowidth.xlsx", "example_noautowidth.xlsx", "example_zip_command.xlsx")
