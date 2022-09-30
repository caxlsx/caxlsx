# encoding: UTF-8
require 'axlsx.rb'

begin
  require "axlsx_styler"

  if defined?(AxlsxStyler)
    raise StandardError.new("Please remove `axlsx_styler` from your Gemfile, the associated functionality is now built-in to `caxlsx` directly.")
  end
rescue LoadError
  # Do nothing, all good
end
