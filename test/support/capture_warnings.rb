module CaptureWarnings
  def capture_warnings
    # Turn off warnings with setting $VERBOSE to nil
    original_verbose = $VERBOSE
    $VERBOSE = nil

    # Redefine warn to redirect warning into our array
    original_warn = Kernel.instance_method(:warn)
    warnings = []
    Kernel.send(:define_method, :warn) { |string| warnings << string }
    yield

    # Revert to the original warn method and set back $VERBOSE to previous value
    Kernel.send(:define_method, :warn, original_warn)
    $VERBOSE = original_verbose

    # Give back the received warnings
    warnings
  end
end
