# frozen_string_literal: true

require "test_helper"
require "open3"
require "tempfile"

class ZipGeneratorTest < Test::Unit::TestCase
  cover Xlsxrb::Ooxml::ZipGenerator
  test "generates a valid zip and preserves entry contents" do
    zip_tempfile = Tempfile.new(["xlsxrb-zip", ".zip"])
    zip_path = zip_tempfile.path
    zip_tempfile.close

    generator = Xlsxrb::Ooxml::ZipGenerator.new(zip_path)
    generator.add_entry("foo.txt", "hello")
    generator.add_entry("nested/bar.txt", "world")
    generator.generate

    stdout, stderr, status = Open3.capture3("unzip", "-t", zip_path)
    assert_equal(true, status.success?, "zip integrity check failed\nSTDOUT:\n#{stdout}\nSTDERR:\n#{stderr}")

    foo_content, foo_err, foo_status = Open3.capture3("unzip", "-p", zip_path, "foo.txt")
    assert_equal(true, foo_status.success?, "failed to read foo.txt\nSTDERR:\n#{foo_err}")
    assert_equal("hello", foo_content)

    bar_content, bar_err, bar_status = Open3.capture3("unzip", "-p", zip_path, "nested/bar.txt")
    assert_equal(true, bar_status.success?, "failed to read nested/bar.txt\nSTDERR:\n#{bar_err}")
    assert_equal("world", bar_content)
  ensure
    File.delete(zip_path) if zip_path && File.exist?(zip_path)
  end

  test "dos_datetime converts Time to MS-DOS packed date and time format" do
    gen = Xlsxrb::Ooxml::ZipGenerator.new("dummy.zip")

    # 2026-09-03 01:15:30
    t = Time.new(2026, 9, 3, 1, 15, 30)
    # Expected: time bytes [0xEF, 0x09], date bytes [0x23, 0x5D]
    assert_equal([239, 9, 35, 93], gen.send(:dos_datetime, t))

    # Boundary: MS-DOS epoch 1980-01-01 00:00:00
    t_epoch = Time.new(1980, 1, 1, 0, 0, 0)
    # date: (0 << 9) | (1 << 5) | 1 = 33 -> [33, 0]
    # time: 0 -> [0, 0]
    assert_equal([0, 0, 33, 0], gen.send(:dos_datetime, t_epoch))

    # Odd seconds are truncated to 2-second resolution (31 sec -> 15)
    t_odd = Time.new(2026, 9, 3, 1, 15, 31)
    assert_equal([239, 9, 35, 93], gen.send(:dos_datetime, t_odd))
  end

  test "crc32, le16, le32, and central directory size calculations" do
    gen = Xlsxrb::Ooxml::ZipGenerator.new("dummy.zip")

    # crc32
    assert_equal(907_060_870, gen.send(:crc32, "hello"))
    assert_equal(0, gen.send(:crc32, ""))

    # le16 and le32 little-endian byte arrays
    assert_equal([0x34, 0x12], gen.send(:le16, 0x1234))
    assert_equal([0x78, 0x56, 0x34, 0x12], gen.send(:le32, 0x12345678))
    assert_equal([0x34, 0x12], Xlsxrb::Ooxml::ZipGenerator.le16(0x1234))
    assert_equal([0x00, 0x00], Xlsxrb::Ooxml::ZipGenerator.le16(0))
    assert_equal([0xFF, 0xFF], Xlsxrb::Ooxml::ZipGenerator.le16(0xFFFF))
    assert_equal([0x78, 0x56, 0x34, 0x12], Xlsxrb::Ooxml::ZipGenerator.le32(0x12345678))
    assert_equal([0x00, 0x00, 0x00, 0x00], Xlsxrb::Ooxml::ZipGenerator.le32(0))
    assert_equal([0xFF, 0xFF, 0xFF, 0xFF], Xlsxrb::Ooxml::ZipGenerator.le32(0xFFFFFFFF))

    # calculate_central_dir_size (46 bytes fixed header + pathname bytesize)
    gen.add_entry("a.txt", "content1")
    gen.add_entry("sub/b.txt", "content2")
    # (46 + 5) + (46 + 9) = 51 + 55 = 106
    assert_equal(106, gen.send(:calculate_central_dir_size))
  end
end
