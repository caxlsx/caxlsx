# frozen_string_literal: true

require 'tc_helper'

class TestView3D < Minitest::Test
  def setup
    @view = Axlsx::View3D.new
  end

  def teardown; end

  def test_options
    v = Axlsx::View3D.new rot_x: 10, rot_y: 5, h_percent: "30%", depth_percent: "45%", r_ang_ax: false, perspective: 10

    assert_equal(10, v.rot_x)
    assert_equal(5, v.rot_y)
    assert_equal("30%", v.h_percent)
    assert_equal("45%", v.depth_percent)
    assert_false(v.r_ang_ax)
    assert_equal(10, v.perspective)
  end

  def test_rot_x
    assert_raises(ArgumentError) { @view.rot_x = "bob" }
    refute_raises { @view.rot_x = -90 }
  end

  def test_rot_y
    assert_raises(ArgumentError) { @view.rot_y = "bob" }
    refute_raises { @view.rot_y = 90 }
  end

  def test_h_percent
    assert_raises(ArgumentError) { @view.h_percent = "bob" }
    refute_raises { @view.h_percent = "500%" }
  end

  def test_depth_percent
    assert_raises(ArgumentError) { @view.depth_percent = "bob" }
    refute_raises { @view.depth_percent = "20%" }
  end

  def test_rAngAx
    assert_raises(ArgumentError) { @view.rAngAx = "bob" }
    refute_raises { @view.rAngAx = true }
  end

  def test_perspective
    assert_raises(ArgumentError) { @view.perspective = "bob" }
    refute_raises { @view.perspective = 30 }
  end
end
