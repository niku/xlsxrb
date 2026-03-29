# frozen_string_literal: true

require "test_helper"
require "open3"
require "tempfile"

class ZipGeneratorTest < Test::Unit::TestCase
  test "generates a valid zip and preserves entry contents" do
    zip_tempfile = Tempfile.new(["xlsxrb-zip", ".zip"])
    zip_path = zip_tempfile.path
    zip_tempfile.close

    generator = Xlsxrb::ZipGenerator.new(zip_path)
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
end
