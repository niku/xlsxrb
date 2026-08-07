# frozen_string_literal: true

require_relative "test_helper"

class ZipBombTest < Test::Unit::TestCase
  test "read mitigates ZIP bombs by limiting uncompressed size" do
    # 500MB of zeros compresses very well
    # We don't actually need to compress 500MB, we can just feed a fake raw deflate stream
    # But it's easier to just mock MAX_UNCOMPRESSED_SIZE to test the logic

    original_limit = Xlsxrb::Ooxml::ZipReader::MAX_UNCOMPRESSED_SIZE
    Xlsxrb::Ooxml::ZipReader.send(:remove_const, :MAX_UNCOMPRESSED_SIZE)
    Xlsxrb::Ooxml::ZipReader::MAX_UNCOMPRESSED_SIZE = 100 # 100 bytes limit

    Tempfile.create(["bomb", ".xlsx"]) do |f|
      # Create a valid zip with one highly compressed file that expands to > 100 bytes
      Xlsxrb.generate(f.path) do |w|
        w.sheet("Sheet1") do |s|
          s.row(["A" * 1000])
        end
      end

      assert_raise(ArgumentError, "ZIP bomb detected") do
        Xlsxrb.read(f.path)
      end
    end
  ensure
    Xlsxrb::Ooxml::ZipReader.send(:remove_const, :MAX_UNCOMPRESSED_SIZE)
    Xlsxrb::Ooxml::ZipReader::MAX_UNCOMPRESSED_SIZE = original_limit
  end
end
