# frozen_string_literal: true

require_relative "test_helper"

class ZipBombTest < Test::Unit::TestCase
  test "read mitigates ZIP bombs by limiting uncompressed size" do
    Tempfile.create(["bomb", ".xlsx"]) do |f|
      # Create a valid zip with one highly compressed file that expands to > 100 bytes
      Xlsxrb.generate(f.path) do |w|
        w.sheet("Sheet1") do |s|
          s.row(["A" * 1000])
        end
      end

      assert_raise(ArgumentError, "ZIP bomb detected") do
        Xlsxrb::Ooxml::ZipReader.open(f.path, max_uncompressed_size: 100, &:read_all)
      end
    end
  end
end
