require 'tc_helper.rb'

class TestAxlsx < Test::Unit::TestCase
  def setup_wide
    @wide_test_points = {
      "A3"    =>                              0,
      "Z3"    =>                             25,
      "B3"    =>                              1,
      "AA3"   =>                  (1 * 26) +  0,
      "AAA3"  => (1 * (26**2)) +  (1 * 26) +  0,
      "AAZ3"  => (1 * (26**2)) +  (1 * 26) + 25,
      "ABA3"  => (1 * (26**2)) +  (2 * 26) +  0,
      "BZU3"  => (2 * (26**2)) + (26 * 26) + 20
    }
  end

  def test_cell_range_empty_if_no_cell
    assert_equal(Axlsx.cell_range([]), "")
  end

  def test_do_not_trust_input_by_default
    assert_equal false, Axlsx.trust_input
  end

  def test_trust_input_can_be_set_to_true
    # Class variables like this are not reset between test runs, so we have
    # to save and restore the original value manually.
    old = Axlsx.trust_input

    Axlsx.trust_input = true
    assert_equal true, Axlsx.trust_input

    Axlsx.trust_input = old
  end

  def test_cell_range_relative
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet
    row = ws.add_row
    c1 = row.add_cell
    c2 = row.add_cell
    assert_equal(Axlsx.cell_range([c2, c1], false), "A1:B1")
  end

  def test_cell_range_absolute
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet :name => "Sheet <'>\" 1"
    row = ws.add_row
    c1 = row.add_cell
    c2 = row.add_cell
    assert_equal(Axlsx.cell_range([c2, c1], true), "'Sheet &lt;''&gt;&quot; 1'!$A$1:$B$1")
  end

  def test_cell_range_row
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet
    row = ws.add_row
    row.add_cell
    row.add_cell
    row.add_cell
    assert_equal("A1:C1", Axlsx.cell_range(row, false))
  end

  def test_name_to_indices
    setup_wide
    @wide_test_points.each do |key, value|
      assert_equal(Axlsx.name_to_indices(key), [value, 2])
    end
  end

  def test_col_ref
    setup_wide
    @wide_test_points.each do |key, value|
      assert_equal(Axlsx.col_ref(value), key.gsub(/\d+/, ''))
    end
  end

  def test_cell_r
    # todo
  end

  def test_range_to_a
    assert_equal([['A1', 'B1', 'C1']],                         Axlsx::range_to_a('A1:C1'))
    assert_equal([['A1', 'B1', 'C1'], ['A2', 'B2', 'C2']],     Axlsx::range_to_a('A1:C2'))
    assert_equal([['Z5', 'AA5', 'AB5'], ['Z6', 'AA6', 'AB6']], Axlsx::range_to_a('Z5:AB6'))
  end

  def test_sanitize_frozen_control_strippped
    needs_sanitize = "legit\x08".freeze # Backspace control char

    assert_equal(Axlsx.sanitize(needs_sanitize), 'legit', 'should strip control chars')
  end

  def test_sanitize_unfrozen_control_strippped
    needs_sanitize = "legit\x08" # Backspace control char
    sanitized_str = Axlsx.sanitize(needs_sanitize)

    assert_equal(sanitized_str,           'legit',                  'should strip control chars')
    assert_equal(sanitized_str.object_id, sanitized_str.object_id,  'should preserve object')
  end

  def test_sanitize_unfrozen_no_sanitize
    legit_str = 'legit'
    sanitized_str = Axlsx.sanitize(legit_str)

    assert_equal(sanitized_str,           legit_str,            'should preserve value')
    assert_equal(sanitized_str.object_id, legit_str.object_id,  'should preserve object')
  end

  class InstanceValuesSubject
    def initialize(args = {})
      args.each do |key, v|
        instance_variable_set("@#{key}".to_sym, v)
      end
    end
  end

  def test_instance_values_for
    empty = InstanceValuesSubject.new
    assert_equal({}, Axlsx.instance_values_for(empty), 'should generate with no ivars')

    single = InstanceValuesSubject.new(a: 2)
    assert_equal({ "a" => 2 }, Axlsx.instance_values_for(single), 'should generate for a single ivar')

    double = InstanceValuesSubject.new(a: 2, b: "c")
    assert_equal({ "a" => 2, "b" => "c" }, Axlsx.instance_values_for(double), 'should generate for multiple ivars')

    inner_obj = Object.new
    complex = InstanceValuesSubject.new(obj: inner_obj)
    assert_equal({ "obj" => inner_obj }, Axlsx.instance_values_for(complex), 'should pass value of ivar directly')

    nil_subject = InstanceValuesSubject.new(nil_obj: nil)
    assert_equal({ "nil_obj" => nil }, Axlsx.instance_values_for(nil_subject), 'should return nil ivars')
  end

  def test_hash_deep_merge
    h1 = { foo: { bar: true } }
    h2 = { foo: { baz: true } }
    assert_equal({ foo: { baz: true } }, h1.merge(h2))
    assert_equal({ foo: { bar: true, baz: true } }, Axlsx.hash_deep_merge(h1, h2))
  end

  def test_escape_formulas
    Axlsx.instance_variable_set(:@escape_formulas, nil)
    assert_equal false, Axlsx::escape_formulas

    Axlsx::escape_formulas = true
    assert_equal true, Axlsx::escape_formulas

    Axlsx::escape_formulas = false
    assert_equal false, Axlsx::escape_formulas
  ensure
    Axlsx.instance_variable_set(:@escape_formulas, nil)
  end
end
