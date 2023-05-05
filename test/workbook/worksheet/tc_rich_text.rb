# frozen_string_literal: true

require 'tc_helper'

class RichText < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name => "hmmmz"
    p.workbook.styles.add_style :sz => 20
    @rt = Axlsx::RichText.new
    b = true
    27.times do |r|
      @rt.add_run "run #{r}, ", :b => (b = !b), :i => !b
    end
    @row = @ws.add_row [@rt]
    @c = @row.first
  end

  def test_initialize
    assert_equal(@c.value, @rt)
    rt_direct = Axlsx::RichText.new('hi', :i => true)
    rt_indirect = Axlsx::RichText.new
    rt_indirect.add_run('hi', :i => true)

    assert_equal(1, rt_direct.runs.length)
    assert_equal(1, rt_indirect.runs.length)
    row = @ws.add_row [rt_direct, rt_indirect]

    assert_equal(row[0].to_xml_string(0, 0), row[1].to_xml_string(0, 0))
  end

  def test_textruns
    runs = @rt.runs

    assert_equal(27, runs.length)
    refute(runs.first.b)
    assert(runs.first.i)
    assert(runs[1].b)
    refute(runs[1].i)
  end

  def test_implicit_richtext
    rt = Axlsx::RichText.new('a', :b => true)
    row_rt = @ws.add_row [rt]
    row_imp = @ws.add_row ['a']
    row_imp[0].b = true

    assert_equal(row_rt[0].to_xml_string(0, 0), row_imp[0].to_xml_string(0, 0))
  end
end
