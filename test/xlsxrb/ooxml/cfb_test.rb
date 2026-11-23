# frozen_string_literal: true

require "test_helper"
require_relative "../../../lib/xlsxrb/ooxml/cfb"

module Xlsxrb
  module Ooxml
    class CfbTest < Test::Unit::TestCase
      cover Xlsxrb::Ooxml::Cfb

      def test_cfb_write_and_read_streams
        streams = {
          "EncryptionInfo" => "Sample Encryption Info XML data <encryption>...</encryption>".b,
          "EncryptedPackage" => ("A" * 5000).b # Spans multiple sectors (> 512 bytes)
        }

        cfb_bytes = Cfb::Writer.write(streams)
        assert_operator cfb_bytes.bytesize, :>, 512
        assert_true Cfb::Reader.cfb?(cfb_bytes)

        reader = Cfb::Reader.new(cfb_bytes)
        assert_equal %w[EncryptionInfo EncryptedPackage], reader.stream_names

        info_read = reader.read_stream("EncryptionInfo")
        assert_equal streams["EncryptionInfo"], info_read

        pkg_read = reader.read_stream("EncryptedPackage")
        assert_equal streams["EncryptedPackage"], pkg_read

        # Non-existent stream
        assert_nil reader.read_stream("NonExistentStream")
      end

      def test_cfb_predicate_and_corrupt_data
        assert_false Cfb::Reader.cfb?(nil)
        assert_false Cfb::Reader.cfb?("")
        assert_false Cfb::Reader.cfb?("short".b)
        assert_false Cfb::Reader.cfb?(123)
        assert_false Cfb::Reader.cfb?(:symbol)
        assert_false Cfb::Reader.cfb?("日本語テキスト".encode(Encoding::UTF_8))
        assert_false Cfb::Reader.cfb?((+"\xFF\xFEabc").force_encoding(Encoding::UTF_8))

        plain_zip = "PK\x03\x04SomeZipData".b
        assert_false Cfb::Reader.cfb?(plain_zip)
        assert_raise(Xlsxrb::Error) do
          Cfb::Reader.new(plain_zip)
        end

        # Header matching magic but too small (< 512 bytes)
        corrupted_short_cfb = Cfb::MAGIC + ("\x00".b * 100)
        assert_raise(Xlsxrb::Error) do
          Cfb::Reader.new(corrupted_short_cfb)
        end

        valid_cfb = Cfb::Writer.write({ "Test" => "Data".b })

        # Corrupted major version
        corrupted_ver = valid_cfb.dup
        corrupted_ver[0x1A, 2] = [99].pack("v")
        assert_raise(Xlsxrb::Error) { Cfb::Reader.new(corrupted_ver) }

        # Corrupted sector shift
        corrupted_shift = valid_cfb.dup
        corrupted_shift[0x1E, 2] = [20].pack("v")
        assert_raise(Xlsxrb::Error) { Cfb::Reader.new(corrupted_shift) }

        # Corrupted mini sector shift
        corrupted_mini_shift = valid_cfb.dup
        corrupted_mini_shift[0x20, 2] = [12].pack("v")
        assert_raise(Xlsxrb::Error) { Cfb::Reader.new(corrupted_mini_shift) }
      end

      # Version 4 files have 4,096-byte sectors, and their header takes the whole first one.
      def test_cfb_reads_version_4_files
        sector_size = 4096
        payload = Array.new(5000) { |i| i % 251 }.pack("C*")
        end_of_chain = 0xFFFFFFFE
        free = 0xFFFFFFFF
        dir_entry = lambda do |name, type, child, start, size|
          utf16 = "#{name}\0".encode("UTF-16LE").b
          utf16.ljust(64, "\0".b) + [utf16.bytesize, type, 1, free, free, child].pack("vCCVVV") +
            ("\0".b * 36) + [start, size].pack("VQ<")
        end

        header = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1".b + ("\0".b * 16) +
                 [0x3E, 4, 0xFFFE, 12, 6].pack("v5") + ("\0".b * 6) +
                 [1, 1, 1, 0, 4096, end_of_chain, 0, end_of_chain, 0].pack("V9") +
                 ([0] + ([free] * 108)).pack("V109")
        # Sector 0 is the FAT, sector 1 the directory, and sectors 2 and 3 the stream.
        fat = [0xFFFFFFFD, end_of_chain, 3, end_of_chain].pack("V4").ljust(sector_size, "\xFF".b)
        directory = (dir_entry.call("Root Entry", Cfb::OBJ_ROOT, 1, end_of_chain, 0) +
                     dir_entry.call("Data", Cfb::OBJ_STREAM, free, 2, payload.bytesize)).ljust(sector_size, "\0".b)
        file = header.ljust(sector_size, "\0".b) + fat + directory + payload.ljust(sector_size * 2, "\0".b)

        reader = Cfb::Reader.new(file)
        assert_equal ["Data"], reader.stream_names
        assert_equal payload, reader.read_stream("Data")

        # The file may end where the last stream ends, without padding the sector, as some
        # files written by Excel 2007 do.
        unpadded = Cfb::Reader.new(file[0, (sector_size * 3) + payload.bytesize])
        assert_equal payload, unpadded.read_stream("Data")

        # A file cut short inside the stream has the bytes before the cut.
        cut = Cfb::Reader.new(file[0, (sector_size * 4) + 100])
        assert_equal payload[0, sector_size + 100], cut.read_stream("Data")
      end

      def test_cfb_dir_entry_predicates
        entry_stream = Cfb::DirEntry.new(name: "TestStream", type: Cfb::OBJ_STREAM)
        assert_true entry_stream.stream?
        assert_false entry_stream.root?
        assert_false entry_stream.storage?

        entry_root = Cfb::DirEntry.new(name: "Root", type: Cfb::OBJ_ROOT)
        assert_false entry_root.stream?
        assert_true entry_root.root?
        assert_false entry_root.storage?

        entry_storage = Cfb::DirEntry.new(name: "Storage", type: Cfb::OBJ_STORAGE)
        assert_false entry_storage.stream?
        assert_false entry_storage.root?
        assert_true entry_storage.storage?
      end
    end
  end
end
