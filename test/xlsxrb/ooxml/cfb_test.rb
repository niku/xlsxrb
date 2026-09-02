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
