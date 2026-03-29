# frozen_string_literal: true

require "test_helper"

class XlsxrbTest < Test::Unit::TestCase
  test "VERSION" do
    assert do
      ::Xlsxrb.const_defined?(:VERSION)
    end
  end
end
