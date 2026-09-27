# frozen_string_literal: true

require "test_helper"

class UtilsTest < Test::Unit::TestCase
  cover Xlsxrb::Ooxml::Utils

  # --- date_to_serial and serial_to_date ---

  test "converts date before, around, and after 1900 leap day bug" do
    # Serial 1 = Jan 1, 1900
    d1 = Date.new(1900, 1, 1)
    assert_equal(1, Xlsxrb::Ooxml::Utils.date_to_serial(d1))
    assert_equal(d1, Xlsxrb::Ooxml::Utils.serial_to_date(1))

    # Serial 59 = Feb 28, 1900 (last day before Lotus 1-2-3 fictitious Feb 29)
    d59 = Date.new(1900, 2, 28)
    assert_equal(59, Xlsxrb::Ooxml::Utils.date_to_serial(d59))
    assert_equal(d59, Xlsxrb::Ooxml::Utils.serial_to_date(59))

    # Serial 61 = Mar 1, 1900 (Lotus 1-2-3 skips 60, raw serial 60 becomes 61)
    d61 = Date.new(1900, 3, 1)
    assert_equal(61, Xlsxrb::Ooxml::Utils.date_to_serial(d61))
    assert_equal(d61, Xlsxrb::Ooxml::Utils.serial_to_date(61))

    # Serial 60 = Feb 29, 1900 (fictitious leap day in Lotus 1-2-3 / Excel)
    # Since 1900 is not a leap year in Gregorian calendar, EPOCH_1900 + 60 equals 1900-03-01.
    assert_equal(Date.new(1900, 3, 1), Xlsxrb::Ooxml::Utils.serial_to_date(60))
    # NOTE: serial_to_date(59) is 1900-02-28, serial_to_date(60) is 1900-03-01, serial_to_date(61) is 1900-03-01
    refute_equal(Xlsxrb::Ooxml::Utils.serial_to_date(59), Xlsxrb::Ooxml::Utils.serial_to_date(60))

    # Modern date round-trip
    d_modern = Date.new(2026, 9, 2)
    serial_modern = Xlsxrb::Ooxml::Utils.date_to_serial(d_modern)
    assert_instance_of(Integer, serial_modern)
    assert_instance_of(Integer, Xlsxrb::Ooxml::Utils.date_to_serial(d1))
    assert_instance_of(Integer, Xlsxrb::Ooxml::Utils.date_to_serial(d59))
    assert_instance_of(Integer, Xlsxrb::Ooxml::Utils.date_to_serial(d61))
    assert_equal(d_modern, Xlsxrb::Ooxml::Utils.serial_to_date(serial_modern))
  end

  # --- datetime_to_serial and serial_to_datetime ---

  test "converts time to fractional serial and round-trips" do
    # Midnight (fraction = 0.0)
    t_midnight = Time.utc(2026, 1, 1, 0, 0, 0)
    serial_midnight = Xlsxrb::Ooxml::Utils.datetime_to_serial(t_midnight)
    assert_equal(0.0, serial_midnight % 1.0)
    assert_equal(t_midnight, Xlsxrb::Ooxml::Utils.serial_to_datetime(serial_midnight))

    # Noon (fraction = 0.5)
    t_noon = Time.utc(2026, 1, 1, 12, 0, 0)
    serial_noon = Xlsxrb::Ooxml::Utils.datetime_to_serial(t_noon)
    assert_in_delta(0.5, serial_noon % 1.0, 1e-6)
    assert_equal(t_noon, Xlsxrb::Ooxml::Utils.serial_to_datetime(serial_noon))

    # 06:00:00 (fraction = 0.25)
    t_morning = Time.utc(2026, 1, 1, 6, 0, 0)
    serial_morning = Xlsxrb::Ooxml::Utils.datetime_to_serial(t_morning)
    assert_in_delta(0.25, serial_morning % 1.0, 1e-6)
    assert_equal(t_morning, Xlsxrb::Ooxml::Utils.serial_to_datetime(serial_morning))

    # Arbitrary time with seconds
    t_detail = Time.utc(2026, 5, 20, 14, 35, 42)
    serial_detail = Xlsxrb::Ooxml::Utils.datetime_to_serial(t_detail)
    assert_equal(t_detail, Xlsxrb::Ooxml::Utils.serial_to_datetime(serial_detail))
  end

  test "converts date in 1904 date system" do
    # Serial 0 = Jan 1, 1904
    d0 = Date.new(1904, 1, 1)
    assert_equal(0, Xlsxrb::Ooxml::Utils.date_to_serial(d0, date1904: true))
    assert_equal(d0, Xlsxrb::Ooxml::Utils.serial_to_date(0, date1904: true))

    # Serial 1 = Jan 2, 1904
    d1 = Date.new(1904, 1, 2)
    assert_equal(1, Xlsxrb::Ooxml::Utils.date_to_serial(d1, date1904: true))
    assert_equal(d1, Xlsxrb::Ooxml::Utils.serial_to_date(1, date1904: true))

    # Serial 59 = Feb 29, 1904 (1904 is a real leap year)
    d59 = Date.new(1904, 2, 29)
    assert_equal(59, Xlsxrb::Ooxml::Utils.date_to_serial(d59, date1904: true))
    assert_equal(d59, Xlsxrb::Ooxml::Utils.serial_to_date(59, date1904: true))

    # Serial 60 = Mar 1, 1904 (no skipped serial in 1904 system)
    d60 = Date.new(1904, 3, 1)
    assert_equal(60, Xlsxrb::Ooxml::Utils.date_to_serial(d60, date1904: true))
    assert_equal(d60, Xlsxrb::Ooxml::Utils.serial_to_date(60, date1904: true))

    # Serial 61 = Mar 2, 1904
    d61 = Date.new(1904, 3, 2)
    assert_equal(61, Xlsxrb::Ooxml::Utils.date_to_serial(d61, date1904: true))
    assert_equal(d61, Xlsxrb::Ooxml::Utils.serial_to_date(61, date1904: true))

    # Difference between 1900 and 1904 system is exactly 1,462 days
    d_modern = Date.new(2026, 9, 27)
    s1900 = Xlsxrb::Ooxml::Utils.date_to_serial(d_modern, date1904: false)
    s1900_default = Xlsxrb::Ooxml::Utils.date_to_serial(d_modern)
    s1904 = Xlsxrb::Ooxml::Utils.date_to_serial(d_modern, date1904: true)
    assert_equal(s1900_default, s1900)
    assert_equal(1462, s1900 - s1904)
    assert_equal(d_modern, Xlsxrb::Ooxml::Utils.serial_to_date(s1904, date1904: true))
    assert_equal(d_modern, Xlsxrb::Ooxml::Utils.serial_to_date(s1900, date1904: false))
    assert_equal(d_modern, Xlsxrb::Ooxml::Utils.serial_to_date(s1900))

    # Pre-1904 boundary
    d_pre = Date.new(1903, 12, 31)
    assert_equal(-1, Xlsxrb::Ooxml::Utils.date_to_serial(d_pre, date1904: true))
    assert_equal(d_pre, Xlsxrb::Ooxml::Utils.serial_to_date(-1, date1904: true))
  end

  test "converts datetime in 1904 date system" do
    # Midnight Jan 1, 1904
    t0 = Time.utc(1904, 1, 1, 0, 0, 0)
    assert_equal(0.0, Xlsxrb::Ooxml::Utils.datetime_to_serial(t0, date1904: true))
    assert_equal(t0, Xlsxrb::Ooxml::Utils.serial_to_datetime(0.0, date1904: true))

    # Noon Jan 1, 1904
    t_noon = Time.utc(1904, 1, 1, 12, 0, 0)
    s_noon = Xlsxrb::Ooxml::Utils.datetime_to_serial(t_noon, date1904: true)
    assert_in_delta(0.5, s_noon, 1e-6)
    assert_equal(t_noon, Xlsxrb::Ooxml::Utils.serial_to_datetime(s_noon, date1904: true))

    # Modern datetime round-trip and 1,462-day offset
    t_modern = Time.utc(2026, 9, 27, 14, 30, 45)
    s1904 = Xlsxrb::Ooxml::Utils.datetime_to_serial(t_modern, date1904: true)
    s1900 = Xlsxrb::Ooxml::Utils.datetime_to_serial(t_modern, date1904: false)
    s1900_default = Xlsxrb::Ooxml::Utils.datetime_to_serial(t_modern)
    assert_equal(s1900_default, s1900)
    assert_in_delta(1462.0, s1900 - s1904, 1e-6)
    assert_equal(t_modern, Xlsxrb::Ooxml::Utils.serial_to_datetime(s1904, date1904: true))
    assert_equal(t_modern, Xlsxrb::Ooxml::Utils.serial_to_datetime(s1900, date1904: false))
    assert_equal(t_modern, Xlsxrb::Ooxml::Utils.serial_to_datetime(s1900))
  end

  # --- hash_password ---

  test "hash_password generates ECMA-376 compliant password hash structure" do
    result = Xlsxrb::Ooxml::Utils.hash_password("secret", spin_count: 10)
    assert_equal("SHA-512", result[:algorithm_name])
    assert_equal(10, result[:spin_count])
    assert_not_nil(result[:hash_value])
    assert_not_nil(result[:salt_value])

    # Default parameters
    default_res = Xlsxrb::Ooxml::Utils.hash_password("secret")
    assert_equal("SHA-512", default_res[:algorithm_name])
    assert_equal(100_000, default_res[:spin_count])
    assert_equal(24, default_res[:salt_value].size) # 16 bytes in Base64 is 24 chars

    # Deterministic with custom salt
    salt = "1234567890123456"
    res1 = Xlsxrb::Ooxml::Utils.hash_password("pass", algorithm: "SHA-256", salt: salt, spin_count: 5)
    res2 = Xlsxrb::Ooxml::Utils.hash_password("pass", algorithm: "SHA-256", salt: salt, spin_count: 5)
    assert_equal(res1[:hash_value], res2[:hash_value])
    assert_equal("SHA-256", res1[:algorithm_name])
    assert_equal(5, res1[:spin_count])

    # Validation errors
    assert_raises(ArgumentError) { Xlsxrb::Ooxml::Utils.hash_password(123) }
    assert_raises(ArgumentError) { Xlsxrb::Ooxml::Utils.hash_password("pass", spin_count: 0) }
    assert_raises(ArgumentError) { Xlsxrb::Ooxml::Utils.hash_password("pass", spin_count: -1) }
    assert_raises(ArgumentError) { Xlsxrb::Ooxml::Utils.hash_password("pass", spin_count: "invalid") }
  end

  # --- Constants and Format Definitions ---

  test "format constants and builtin tables are complete and frozen" do
    assert_equal([14, 15, 16, 17, 18, 19, 20, 21, 22], Xlsxrb::Ooxml::Utils::BUILTIN_DATE_FMT_IDS)
    assert(Xlsxrb::Ooxml::Utils::BUILTIN_DATE_FMT_IDS.frozen?)

    assert_equal("yyyy\\-mm\\-dd", Xlsxrb::Ooxml::Utils::DEFAULT_DATE_FORMAT)
    assert_equal("yyyy\\-mm\\-dd\\ hh:mm:ss", Xlsxrb::Ooxml::Utils::DEFAULT_DATETIME_FORMAT)

    assert_equal("General", Xlsxrb::Ooxml::Utils::BUILTIN_NUM_FMT_CODES[0])
    assert_equal("mm-dd-yy", Xlsxrb::Ooxml::Utils::BUILTIN_NUM_FMT_CODES[14])
    assert_equal("@", Xlsxrb::Ooxml::Utils::BUILTIN_NUM_FMT_CODES[49])
    assert(Xlsxrb::Ooxml::Utils::BUILTIN_NUM_FMT_CODES.frozen?)
  end
end
