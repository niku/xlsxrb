# frozen_string_literal: true

require_relative "test_helper"

class MagicNumberTest < Test::Unit::TestCase
  test "read raises ArgumentError for invalid magic number" do
    Tempfile.create(["invalid", ".xlsx"]) do |f|
      f.write "This is not a zip file"
      f.close

      assert_raise(ArgumentError, "Invalid magic number") do
        Xlsxrb.read(f.path)
      end
    end
  end
end
