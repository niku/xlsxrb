# frozen_string_literal: true

require "test_helper"

class EncryptionSecurityAuditTest < Test::Unit::TestCase
  # 1. Verification of Underlying Crypto Primitives (OS OpenSSL C-library)
  def test_crypto_engine_uses_native_openssl_primitives
    assert_not_nil defined?(OpenSSL::Cipher)
    assert_not_nil defined?(OpenSSL::Digest)
    assert_not_nil defined?(SecureRandom)

    # Ensure CSPRNG is cryptographically random (not predictable pseudo-random)
    salts = Array.new(100) { SecureRandom.random_bytes(16) }
    assert_equal 100, salts.uniq.size, "CSPRNG produced duplicate salts!"
  end

  # 2. Invalid Password Rejection
  def test_invalid_password_rejection_for_various_mutations
    Dir.mktmpdir do |dir|
      xlsx_path = File.join(dir, "secure_timing.xlsx")
      Xlsxrb.write(xlsx_path, password: "CorrectSecretPassword2026") do |wb|
        wb.sheet("Data") { |s| s.row(["Private Info", 123]) }
      end

      # Attempt with slightly modified passwords of varying lengths
      bad_passwords = [
        "CorrectSecretPassword2025", # last char different
        "DorrectSecretPassword2026", # first char different
        "WrongPassword" # totally different
      ]

      bad_passwords.each do |bad_pw|
        assert_raise(Xlsxrb::InvalidPasswordError) do
          Xlsxrb.read(xlsx_path, password: bad_pw).load
        end
      end
    end
  end

  # 3. Timing Attack Mitigation (Verifying OpenSSL.secure_compare is explicitly invoked)
  def test_password_verification_uses_constant_time_secure_compare
    calls = 0
    original_method = OpenSSL.method(:secure_compare)

    # Silence Ruby method redefinition warning during test spying
    prev_verbose = $VERBOSE
    $VERBOSE = nil
    OpenSSL.define_singleton_method(:secure_compare) do |a, b|
      calls += 1
      original_method.call(a, b)
    end
    $VERBOSE = prev_verbose

    begin
      Dir.mktmpdir do |dir|
        # Test Standard Mode verification invokes secure_compare
        std_path = File.join(dir, "std.xlsx")
        Xlsxrb.write(std_path, password: "Pass", encryption_mode: :standard) do |wb|
          wb.sheet("S") { |s| s.row([1]) }
        end
        Xlsxrb.read(std_path, password: "Pass").load
        assert_operator calls, :>=, 1

        # Test Agile Mode verification invokes secure_compare
        calls_before_agile = calls
        agile_path = File.join(dir, "agile.xlsx")
        Xlsxrb.write(agile_path, password: "Pass", encryption_mode: :agile) do |wb|
          wb.sheet("S") { |s| s.row([1]) }
        end
        Xlsxrb.read(agile_path, password: "Pass").load
        assert_operator calls, :>, calls_before_agile
      end
    ensure
      prev_verbose = $VERBOSE
      $VERBOSE = nil
      OpenSSL.define_singleton_method(:secure_compare, original_method)
      $VERBOSE = prev_verbose
    end
  end

  # 4. DoS Protection: Memory Exhaustion / Huge total_size Spoofing
  def test_memory_exhaustion_dos_protection_on_huge_total_size
    Dir.mktmpdir do |dir|
      xlsx_path = File.join(dir, "huge_size_attack.xlsx")
      Xlsxrb.write(xlsx_path, password: "ValidPassword123") do |wb|
        wb.sheet("Data") { |s| s.row(["Value", 1]) }
      end

      raw_cfb = File.binread(xlsx_path)
      reader = Xlsxrb::Ooxml::Cfb::Reader.new(raw_cfb)
      enc_pkg = reader.read_stream("EncryptedPackage")
      enc_info = reader.read_stream("EncryptionInfo")

      # Forge an EncryptedPackage with a malicious 16 Exabyte size header (0x7FFFFFFFFFFFFFFF)
      malicious_pkg = [0x7FFF_FFFF_FFFF_FFFF].pack("Q<") + enc_pkg[8..]

      # Rebuild malicious CFB
      corrupted_cfb = Xlsxrb::Ooxml::Cfb::Writer.write(
        "EncryptionInfo" => enc_info,
        "EncryptedPackage" => malicious_pkg
      )

      # Invariant: Must safely reject with DecryptionError without allocating huge memory
      assert_raise(Xlsxrb::DecryptionError) do
        Xlsxrb.read(corrupted_cfb, password: "ValidPassword123").load
      end
    end
  end

  # 5. DoS Protection: Circular Sector Chain in CFB (Infinite Loop Prevention)
  def test_circular_sector_chain_infinite_loop_protection
    # Create synthetic CFB data with circular sector chain (sector 0 -> 1 -> 0)
    # Cfb::Reader must not hang and safely terminate reading
    header = +""
    header << Xlsxrb::Ooxml::Cfb::MAGIC
    header << ("\x00".b * 16)
    header << [0x003B, 0x0003, 0xFFFE, 9, 6].pack("v5")
    header << ("\x00".b * 6)
    header << [0, 1, 0, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE, 0].pack("V9")
    difat = Array.new(109, Xlsxrb::Ooxml::Cfb::FREESECT)
    difat[0] = 2 # Sector 2 is FAT
    header << difat.pack("V109")

    # Sector 0: Stream Data chunk 1
    sec0 = ("A" * 512).b
    # Sector 1: Stream Data chunk 2
    sec1 = ("B" * 512).b
    # Sector 2: FAT table with circular chain: 0 -> 1 -> 0
    fat = Array.new(128, Xlsxrb::Ooxml::Cfb::FREESECT)
    fat[0] = 1
    fat[1] = 0 # Circular!
    fat[2] = Xlsxrb::Ooxml::Cfb::FATSECT
    sec2 = fat.pack("V*")

    raw_cfb = header + sec0 + sec1 + sec2
    assert_true Xlsxrb::Ooxml::Cfb::Reader.cfb?(raw_cfb)

    # Invariant: Initializing reader or reading regular stream must safely terminate without infinite loop
    reader = Xlsxrb::Ooxml::Cfb::Reader.new(raw_cfb)
    data = reader.send(:read_regular_stream_data, 0, 100_000)
    assert_operator data.bytesize, :<=, 1024

    # High-level API invariant: Calling Xlsxrb.read on circular CFB must fail safely with DecryptionError without infinite loop
    assert_raise(Xlsxrb::DecryptionError) do
      Xlsxrb.read(raw_cfb, password: "Pass").load
    end

    # Invariant: Mini-stream circular loop must also terminate safely
    reader.instance_variable_set(:@mini_stream, ("M" * 256).b)
    reader.instance_variable_set(:@minifat, [1, 0]) # Circular mini-fat
    mini_data = reader.send(:read_mini_stream_data, 0, 100_000)
    assert_operator mini_data.bytesize, :<=, 128
  end

  # 6. Nonce / Salt Replay Attack Resistance (CSPRNG Uniqueness across Standard & Agile)
  def test_each_encryption_generates_unique_keys_and_ivs
    plain = "PK\x03\x04Dummy payload content"
    password = "SamePassword123"

    # Test Standard Mode uniqueness
    std1 = Xlsxrb::Ooxml::Crypto.encrypt(plain, password, mode: :standard)
    std2 = Xlsxrb::Ooxml::Crypto.encrypt(plain, password, mode: :standard)
    assert_not_equal std1, std2
    assert_equal plain, Xlsxrb::Ooxml::Crypto.decrypt(std1, password)
    assert_equal plain, Xlsxrb::Ooxml::Crypto.decrypt(std2, password)

    # Test Agile Mode uniqueness
    agile1 = Xlsxrb::Ooxml::Crypto.encrypt(plain, password, mode: :agile)
    agile2 = Xlsxrb::Ooxml::Crypto.encrypt(plain, password, mode: :agile)
    assert_not_equal agile1, agile2
    assert_equal plain, Xlsxrb::Ooxml::Crypto.decrypt(agile1, password)
    assert_equal plain, Xlsxrb::Ooxml::Crypto.decrypt(agile2, password)
  end

  # 7. DoS Protection: Excessive spinCount KDF Amplification Attack
  def test_agile_excessive_spincount_dos_rejection
    # Synthetic Agile EncryptionInfo with malicious spinCount = 999,999,999
    malicious_xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
        <keyData saltValue="AAAA" hashAlgorithm="SHA512" cipherAlgorithm="AES" keyBits="256" blockSize="16"/>
        <keyEncryptors>
          <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
            <encryptedKey spinCount="999999999" saltValue="AAAA" hashAlgorithm="SHA512" cipherAlgorithm="AES" keyBits="256" hashSize="64" encryptedVerifierHashInput="AAAA" encryptedVerifierHashValue="AAAA" encryptedKeyValue="AAAA"/>
          </keyEncryptor>
        </keyEncryptors>
      </encryption>
    XML

    malicious_info = [4, 4, 0x40].pack("vvV") + malicious_xml.b
    dummy_pkg = [100].pack("Q<") + ("A" * 128).b

    corrupted_cfb = Xlsxrb::Ooxml::Cfb::Writer.write(
      "EncryptionInfo" => malicious_info,
      "EncryptedPackage" => dummy_pkg
    )

    # Invariant: Must immediately reject with DecryptionError without spinning CPU 1 billion times
    assert_raise(Xlsxrb::DecryptionError) do
      Xlsxrb.read(corrupted_cfb, password: "Pass").load
    end
  end

  # 8. Malformed / Broken XML Handling in Agile EncryptionInfo
  def test_agile_malformed_xml_rejection
    bad_info = [4, 4, 0x40].pack("vvV") + "<unclosed_tag><broken>".b
    dummy_pkg = [100].pack("Q<") + ("A" * 128).b

    corrupted_cfb = Xlsxrb::Ooxml::Cfb::Writer.write(
      "EncryptionInfo" => bad_info,
      "EncryptedPackage" => dummy_pkg
    )

    assert_raise(Xlsxrb::DecryptionError) do
      Xlsxrb.read(corrupted_cfb, password: "Pass").load
    end
  end

  # 9. Out-of-bounds Header Size in Standard EncryptionInfo
  def test_standard_out_of_bounds_header_size_rejection
    # header_size = 0x7FFFFFFF (causes verifier_offset to point beyond stream length)
    bad_info = [3, 2, 0x24, 0x7FFF_FFFF].pack("vvVV") + ("A" * 50).b
    dummy_pkg = [100].pack("Q<") + ("A" * 128).b

    corrupted_cfb = Xlsxrb::Ooxml::Cfb::Writer.write(
      "EncryptionInfo" => bad_info,
      "EncryptedPackage" => dummy_pkg
    )

    assert_raise(Xlsxrb::DecryptionError) do
      Xlsxrb.read(corrupted_cfb, password: "Pass").load
    end
  end
end
