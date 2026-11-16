# frozen_string_literal: true

require_relative "test_helper"

class ZipBombTest < Test::Unit::TestCase
  test "read mitigates ZIP bombs by limiting uncompressed size" do
    Tempfile.create(["bomb", ".xlsx"]) do |f|
      # Create a valid zip with one highly compressed file that expands to > 100 bytes
      Xlsxrb.write(f.path) do |w|
        w.sheet("Sheet1") do |s|
          s.row(["A" * 1000])
        end
      end

      assert_raise(Xlsxrb::ZipBombError) do
        Xlsxrb::Ooxml::ZipReader.open(f.path, max_uncompressed_size: 100, &:read_all)
      end
    end
  end

  test "streaming read (Xlsxrb.read block form) enforces max_uncompressed_size limit" do
    Tempfile.create(["bomb_streaming", ".xlsx"]) do |f|
      Xlsxrb.write(f.path) do |w|
        w.sheet("Sheet1") do |s|
          s.row(["B" * 2000])
        end
      end

      assert_raise(Xlsxrb::ZipBombError) do
        Xlsxrb.read(f.path, max_uncompressed_size: 100) do |sheet|
          sheet.each { |row| assert_not_nil row }
        end
      end
    end
  end

  test "ZipBombError is a subclass of Xlsxrb::ZipError and Xlsxrb::Error" do
    assert_operator Xlsxrb::ZipBombError, :<, Xlsxrb::ZipError
    assert_operator Xlsxrb::ZipBombError, :<, Xlsxrb::Error
  end

  test "an archive without a central directory does not hold every inflated entry in memory" do
    # Three data-descriptor entries of 1 MB each. Finding where each one ends means inflating
    # it; keeping the output would hold 3 MB for a limit of 2 MB, and without bound for more
    # entries.
    archive = +"".b
    3.times { |i| archive << data_descriptor_entry("xl/part#{i}.bin", "\0" * (1024 * 1024)) }

    megabyte_strings = -> { ObjectSpace.each_object(String).select { |str| str.bytesize >= 1024 * 1024 }.map(&:object_id) }
    before = megabyte_strings.call

    Xlsxrb::Ooxml::ZipReader.open(archive, max_uncompressed_size: 2 * 1024 * 1024) do |zip|
      assert_equal %w[xl/part0.bin xl/part1.bin xl/part2.bin], zip.entry_names
      GC.start
      assert_empty megabyte_strings.call - before
      assert_equal 1024 * 1024, zip.read_entry("xl/part2.bin").bytesize
      assert_equal 1024 * 1024, zip.entry_catalog["xl/part2.bin"].uncompressed_size
    end
  end

  test "a data-descriptor entry over the limit raises ZipBombError while the archive is scanned" do
    archive = data_descriptor_entry("xl/part.bin", "\0" * 4096)

    assert_raise(Xlsxrb::ZipBombError) do
      Xlsxrb::Ooxml::ZipReader.open(archive, max_uncompressed_size: 100, &:entry_names)
    end
  end

  test "an entry whose declared size runs past the end of the archive raises ZipError" do
    name = "xl/workbook.xml"
    # 2**62 comes from a ZIP64 extra field; IO#read would try to allocate it.
    { 0xFFFF_FFF0 => "".b, 0xFFFF_FFFF => [1, 8, 2**62].pack("vvQ<") }.each do |declared_size, extra|
      local = local_entry(name, "x")
      central = ["PK\x01\x02".b, [45, 45, 0, 8, 0, 33, 0, declared_size, 1, name.bytesize, extra.bytesize, 0, 0, 0, 0, 0].pack("vvvvvvVVVvvvvvVV"), name, extra].join
      eocd = ["PK\x05\x06".b, [0, 0, 1, 1, central.bytesize, local.bytesize, 0].pack("vvvvVVv")].join

      Tempfile.create(["forged_size", ".xlsx"]) do |f|
        f.binmode
        f.write(local + central + eocd)
        f.flush

        Xlsxrb::Ooxml::ZipReader.open(f.path) do |zip|
          error = assert_raise(Xlsxrb::ZipError) { zip.read_entry(name) }
          assert_match(/extends past the end of the archive/, error.message)
          assert_raise(Xlsxrb::ZipError) { zip.open_entry_io(name) }
          assert_raise(Xlsxrb::ZipError) { zip.each_entry_chunk(name) { |chunk| chunk } }
        end
      end
    end
  end

  test "malformed archives raise ZipError instead of Zlib or NoMethodError" do
    name = "xl/workbook.xml"

    unsupported = local_entry(name, "hello", method: 12)
    Xlsxrb::Ooxml::ZipReader.open(unsupported) do |zip|
      assert_raise(Xlsxrb::ZipError) { zip.read_entry(name) }
      assert_raise(Xlsxrb::ZipError) { zip.open_entry_io(name).read }
    end

    corrupt = local_entry(name, "hello", compressed: "\xFF\xFF\xFF\xFF".b)
    Xlsxrb::Ooxml::ZipReader.open(corrupt) do |zip|
      assert_raise(Xlsxrb::ZipError) { zip.read_entry(name) }
      assert_raise(Xlsxrb::ZipError) { zip.open_entry_io(name).read }
    end

    # The header declares a 500-byte name and the archive ends after the header.
    truncated = "PK\x03\x04".b + [20, 0, 8, 0, 33, 0, 5, 5, 500, 0].pack("vvvvvVVVvv")
    assert_equal [], Xlsxrb::Ooxml::ZipReader.open(truncated, &:entry_names)
    assert_raise(Xlsxrb::ParseError) { Xlsxrb.load(truncated) }
  end

  private

  def raw_deflate(data)
    deflater = Zlib::Deflate.new(Zlib::BEST_COMPRESSION, -Zlib::MAX_WBITS)
    deflater.deflate(data, Zlib::FINISH)
  ensure
    deflater.close
  end

  def local_entry(name, data, method: 8, compressed: nil)
    compressed ||= method == 8 ? raw_deflate(data) : data
    header = [20, 0, method, 0, 33, Zlib.crc32(data), compressed.bytesize, data.bytesize, name.bytesize, 0].pack("vvvvvVVVvv")
    ["PK\x03\x04".b, header, name.b, compressed].join
  end

  # Local header with sizes of 0 and bit 3 set, followed by the data and a data descriptor.
  def data_descriptor_entry(name, data)
    compressed = raw_deflate(data)
    header = [20, 8, 8, 0, 33, 0, 0, 0, name.bytesize, 0].pack("vvvvvVVVvv")
    descriptor = ["PK\x07\x08".b, [Zlib.crc32(data), compressed.bytesize, data.bytesize].pack("VVV")].join
    ["PK\x03\x04".b, header, name.b, compressed, descriptor].join
  end
end
