# frozen_string_literal: true

require "test_helper"

class NumberFormatterTest < Test::Unit::TestCase
  test "general and default formatting" do
    assert_equal("", Xlsxrb::NumberFormatter.format(nil))
    assert_equal("1042", Xlsxrb::NumberFormatter.format(1042, "General"))
    assert_equal("1042", Xlsxrb::NumberFormatter.format(1042, "GENERAL"))
    assert_equal("1042", Xlsxrb::NumberFormatter.format(1042.0, "General"))
    assert_equal("1042.5", Xlsxrb::NumberFormatter.format(1042.5, "General"))
    assert_equal("-1042", Xlsxrb::NumberFormatter.format(-1042, "General"))
    assert_equal("Hello", Xlsxrb::NumberFormatter.format("Hello", "General"))
    assert_equal("1042", Xlsxrb.format(1042))
  end

  test "boolean formatting" do
    assert_equal("TRUE", Xlsxrb::NumberFormatter.format(true))
    assert_equal("FALSE", Xlsxrb::NumberFormatter.format(false))
  end

  test "cell error formatting" do
    assert_equal("#DIV/0!", Xlsxrb::NumberFormatter.format(Xlsxrb::Error::DIV0))
    assert_equal("#N/A", Xlsxrb::NumberFormatter.format("#N/A"))
    assert_equal("#VALUE!", Xlsxrb::NumberFormatter.format("#VALUE!"))
  end

  test "integer and zero padding" do
    assert_equal("1042", Xlsxrb::NumberFormatter.format(1042, "0"))
    assert_equal("001042", Xlsxrb::NumberFormatter.format(1042, "000000"))
    assert_equal("-001042", Xlsxrb::NumberFormatter.format(-1042, "000000"))
    assert_equal("05010", Xlsxrb::NumberFormatter.format(5010, "00000"))
  end

  test "fixed-point decimals and thousand separators" do
    assert_equal("1042.00", Xlsxrb::NumberFormatter.format(1042, "0.00"))
    assert_equal("1042.0000", Xlsxrb::NumberFormatter.format(1042, "0.0000"))
    assert_equal("1042.000000000", Xlsxrb::NumberFormatter.format(1042, "0.000000000"))
    assert_equal("1,042", Xlsxrb::NumberFormatter.format(1042, "#,##0"))
    assert_equal("1,042.00", Xlsxrb::NumberFormatter.format(1042, "#,##0.00"))
    assert_equal("1,042.000", Xlsxrb::NumberFormatter.format(1042, "#,##0.000"))
    assert_equal("-1,042.50", Xlsxrb::NumberFormatter.format(-1042.5, "#,##0.00"))
  end

  test "percentages" do
    assert_equal("104200%", Xlsxrb::NumberFormatter.format(1042, "0%"))
    assert_equal("57%", Xlsxrb::NumberFormatter.format(0.569999999995, "0%"))
    assert_equal("12.50%", Xlsxrb::NumberFormatter.format(0.125, "0.00%"))
    assert_equal("104200.00%", Xlsxrb::NumberFormatter.format(1042, "0.00%"))
  end

  test "scientific notation" do
    assert_equal("1.04E+03", Xlsxrb::NumberFormatter.format(1042, "0.00E+00"))
    assert_equal("1.0E+03", Xlsxrb::NumberFormatter.format(1042, "##0.0E+0"))
  end

  test "negative patterns with sections" do
    assert_equal("1,042", Xlsxrb::NumberFormatter.format(1042, "#,##0 ;(#,##0)"))
    assert_equal("(1,042)", Xlsxrb::NumberFormatter.format(-1042, "#,##0 ;(#,##0)"))

    assert_equal("1,042", Xlsxrb::NumberFormatter.format(1042, "#,##0 ;[Red](#,##0)"))
    assert_equal("[Red](1,042)", Xlsxrb::NumberFormatter.format(-1042, "#,##0 ;[Red](#,##0)"))

    assert_equal("1,042.00", Xlsxrb::NumberFormatter.format(1042, "#,##0.00;(#,##0.00)"))
    assert_equal("(1,042.00)", Xlsxrb::NumberFormatter.format(-1042, "#,##0.00;(#,##0.00)"))

    assert_equal("1,042.00", Xlsxrb::NumberFormatter.format(1042, "#,##0.00;[Red](#,##0.00)"))
    assert_equal("[Red](1,042.00)", Xlsxrb::NumberFormatter.format(-1042, "#,##0.00;[Red](#,##0.00)"))

    acct = "_-* #,##0.00\\ _€_-;\\-* #,##0.00\\ _€_-;_-* \"-\"??\\ _€_-;_-@_-"
    assert_equal("1,042.00", Xlsxrb::NumberFormatter.format(1042, acct))
    assert_equal("-1,042.00", Xlsxrb::NumberFormatter.format(-1042, acct))
  end

  test "currencies" do
    assert_equal("$20.51", Xlsxrb::NumberFormatter.format(20.51, "\"$\"#,##0.00"))
    assert_equal("$20.51", Xlsxrb::NumberFormatter.format(20.51, "$#,##0.00"))
    assert_equal("$20.51", Xlsxrb::NumberFormatter.format(20.51, "[$$-409]#,##0.00"))
    assert_equal("€20.51", Xlsxrb::NumberFormatter.format(20.51, "[$€-2] #,##0.00"))
    assert_equal("£20.51", Xlsxrb::NumberFormatter.format(20.51, "[$£-809]#,##0.00"))
    assert_equal("¥1,042", Xlsxrb::NumberFormatter.format(1042, "[$¥-411]#,##0"))
    assert_equal("¥1,042", Xlsxrb::NumberFormatter.format(1042, "¥#,##0"))
  end

  test "text formats" do
    assert_equal("1042", Xlsxrb::NumberFormatter.format(1042, "@"))
    assert_equal("Sample", Xlsxrb::NumberFormatter.format("Sample", "@"))
    assert_equal("Value: Sample", Xlsxrb::NumberFormatter.format("Sample", "Value: @"))
  end

  test "date and datetime formats" do
    dt = DateTime.new(2015, 1, 25, 8, 15, 0)
    assert_equal("01-25-15", Xlsxrb::NumberFormatter.format(dt, "mm-dd-yy"))
    assert_equal("25-JAN-15", Xlsxrb::NumberFormatter.format(dt, "d-mmm-yy"))
    assert_equal("25-JAN", Xlsxrb::NumberFormatter.format(dt, "d-mmm"))
    assert_equal("JAN-15", Xlsxrb::NumberFormatter.format(dt, "mmm-yy"))
    assert_equal("1/25/15 8:15", Xlsxrb::NumberFormatter.format(dt, "m/d/yy h:mm"))
    assert_equal("8:15", Xlsxrb::NumberFormatter.format(dt, "h:mm"))
    assert_equal("8:15:00", Xlsxrb::NumberFormatter.format(dt, "h:mm:ss"))
    assert_equal("8:15 AM", Xlsxrb::NumberFormatter.format(dt, "h:mm AM/PM"))
    assert_equal("2015/01/25 08:15:00", Xlsxrb::NumberFormatter.format(dt, "yyyy/mm/dd hh:mm:ss"))
    assert_equal("2015-01-25", Xlsxrb::NumberFormatter.format(dt, "yyyy-mm-dd"))

    d = Date.new(2014, 6, 1)
    assert_equal("06-01-14", Xlsxrb::NumberFormatter.format(d, "mm-dd-yy"))
    assert_equal("2014-06-01", Xlsxrb::NumberFormatter.format(d, "yyyy-mm-dd"))
  end

  test "numeric serial date conversions" do
    assert_equal("06-01-14", Xlsxrb::NumberFormatter.format(41_791, "mm-dd-yy"))
    assert_equal("2014-06-01", Xlsxrb::NumberFormatter.format(41_791, "yyyy-mm-dd"))
    assert_equal("06-02-18", Xlsxrb::NumberFormatter.format(41_791, "mm-dd-yy", date1904: true))

    # Time serial 0.0751 (1:48:09)
    assert_equal("1:48", Xlsxrb::NumberFormatter.format(0.0751, "h:mm"))
    assert_equal("1:48:09", Xlsxrb::NumberFormatter.format(0.0751, "h:mm:ss"))
    assert_equal("48:09", Xlsxrb::NumberFormatter.format(0.0751, "mm:ss"))
    assert_equal("[1]:48:09", Xlsxrb::NumberFormatter.format(0.0751, "[h]:mm:ss"))
  end

  test "format_code_for resolution" do
    styles = {
      cell_xfs: [
        { num_fmt_id: 0 },
        { num_fmt_id: 14 },
        { num_fmt_id: 165 },
        { num_fmt_id: 0, xf_id: 0 }
      ],
      cell_style_xfs: [
        { num_fmt_id: 9 }
      ],
      num_fmts: {
        165 => "$#,##0.00"
      }
    }

    assert_equal("General", Xlsxrb::NumberFormatter.format_code_for(0, styles))
    assert_equal("mm-dd-yy", Xlsxrb::NumberFormatter.format_code_for(1, styles))
    assert_equal("$#,##0.00", Xlsxrb::NumberFormatter.format_code_for(2, styles))
    assert_equal("0%", Xlsxrb::NumberFormatter.format_code_for(3, styles))
    assert_nil(Xlsxrb::NumberFormatter.format_code_for(nil, styles))
    assert_nil(Xlsxrb::NumberFormatter.format_code_for(10, styles))
    assert_nil(Xlsxrb::NumberFormatter.format_code_for(0, nil))
  end

  test "format_type classification and boolean predicates" do
    assert_nil(Xlsxrb::NumberFormatter.format_type(nil))
    assert_equal(:general, Xlsxrb::NumberFormatter.format_type(""))
    assert_equal(:general, Xlsxrb::NumberFormatter.format_type("General"))
    assert_equal(:text, Xlsxrb::NumberFormatter.format_type("@"))

    # Dates
    %w[mm-dd-yy d-mmm-yy d-mmm mmm-yy yyyy-mm-dd yyyy\\-mm\\-dd m/d/yyyy [$-409]dd-mmm-yyyy].each do |code|
      assert_equal(:date, Xlsxrb::NumberFormatter.format_type(code), "Expected :date for #{code}")
      assert_true(Xlsxrb::NumberFormatter.date_format?(code))
      assert_true(Xlsxrb::NumberFormatter.date_only_format?(code))
      assert_false(Xlsxrb::NumberFormatter.datetime_format?(code))
      assert_false(Xlsxrb::NumberFormatter.time_format?(code))
    end

    # Datetimes
    ["m/d/yy h:mm", "yyyy-mm-dd hh:mm:ss", "yyyy\\-mm\\-dd\\ hh:mm:ss"].each do |code|
      assert_equal(:datetime, Xlsxrb::NumberFormatter.format_type(code), "Expected :datetime for #{code}")
      assert_true(Xlsxrb::NumberFormatter.date_format?(code))
      assert_false(Xlsxrb::NumberFormatter.date_only_format?(code))
      assert_true(Xlsxrb::NumberFormatter.datetime_format?(code))
      assert_false(Xlsxrb::NumberFormatter.time_format?(code))
    end

    # Times
    ["h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "mm:ss", "[h]:mm:ss", "mmss.0"].each do |code|
      assert_equal(:time, Xlsxrb::NumberFormatter.format_type(code), "Expected :time for #{code}")
      assert_true(Xlsxrb::NumberFormatter.date_format?(code))
      assert_false(Xlsxrb::NumberFormatter.date_only_format?(code))
      assert_false(Xlsxrb::NumberFormatter.datetime_format?(code))
      assert_true(Xlsxrb::NumberFormatter.time_format?(code))
    end

    # Numbers
    ["0", "0.00", "#,##0.00", "0%", "0.00E+00", "$#,##0.00", "[Red]#,##0"].each do |code|
      assert_equal(:number, Xlsxrb::NumberFormatter.format_type(code), "Expected :number for #{code}")
      assert_false(Xlsxrb::NumberFormatter.date_format?(code))
      assert_false(Xlsxrb::NumberFormatter.datetime_format?(code))
      assert_false(Xlsxrb::NumberFormatter.time_format?(code))
    end
  end

  test "Elements::Styles precomputed number format classifications" do
    raw_styles = {
      cell_xfs: [
        { num_fmt_id: 0 },
        { num_fmt_id: 14 },
        { num_fmt_id: 20 },
        { num_fmt_id: 22 },
        { num_fmt_id: 165 },
        { num_fmt_id: 166 },
        { num_fmt_id: 0, xf_id: 0 }
      ],
      cell_style_xfs: [
        { num_fmt_id: 15 }
      ],
      num_fmts: {
        165 => "yyyy\\-mm\\-dd",
        166 => "[h]:mm:ss"
      }
    }

    styles = Xlsxrb::Elements::Styles.new(raw_styles)
    styles.precompute!

    assert_true(styles.is_a?(Hash))
    assert_equal(7, styles[:cell_xfs].size)

    # Style 0: General
    assert_equal("General", styles.number_format(0))
    assert_equal(:general, styles.format_type(0))
    assert_false(styles.date_format?(0))

    # Style 1: Builtin 14 (mm-dd-yy)
    assert_equal("mm-dd-yy", styles.number_format(1))
    assert_equal(:date, styles.format_type(1))
    assert_true(styles.date_format?(1))
    assert_true(styles.date_only_format?(1))
    assert_false(styles.datetime_format?(1))
    assert_false(styles.time_format?(1))

    # Style 2: Builtin 20 (h:mm)
    assert_equal("h:mm", styles.number_format(2))
    assert_equal(:time, styles.format_type(2))
    assert_true(styles.date_format?(2))
    assert_false(styles.date_only_format?(2))
    assert_false(styles.datetime_format?(2))
    assert_true(styles.time_format?(2))

    # Style 3: Builtin 22 (m/d/yy h:mm)
    assert_equal("m/d/yy h:mm", styles.number_format(3))
    assert_equal(:datetime, styles.format_type(3))
    assert_true(styles.date_format?(3))
    assert_false(styles.date_only_format?(3))
    assert_true(styles.datetime_format?(3))
    assert_false(styles.time_format?(3))

    # Style 4: Custom Date (yyyy-mm-dd)
    assert_equal("yyyy\\-mm\\-dd", styles.number_format(4))
    assert_equal(:date, styles.format_type(4))
    assert_true(styles.date_format?(4))
    assert_true(styles.date_only_format?(4))

    # Style 5: Custom Time ([h]:mm:ss)
    assert_equal("[h]:mm:ss", styles.number_format(5))
    assert_equal(:time, styles.format_type(5))
    assert_true(styles.date_format?(5))
    assert_true(styles.time_format?(5))

    # Style 6: Inherited from cell_style_xfs (Builtin 15 = d-mmm-yy)
    assert_equal("d-mmm-yy", styles.number_format(6))
    assert_equal(:date, styles.format_type(6))
    assert_true(styles.date_format?(6))

    # Out-of-bounds & invalid style indices
    assert_nil(styles.number_format(999))
    assert_nil(styles.format_type(999))
    assert_false(styles.date_format?(999))
    assert_false(styles.datetime_format?(999))
    assert_false(styles.time_format?(999))
    assert_nil(styles.number_format(nil))
    assert_false(styles.date_format?(nil))
    assert_nil(styles.number_format(-1))
    assert_false(styles.date_format?(-1))

    # NumberFormatter delegates to Elements::Styles methods
    assert_equal("mm-dd-yy", Xlsxrb::NumberFormatter.format_code_for(1, styles))
    assert_equal(:date, Xlsxrb::NumberFormatter.format_type_for(1, styles))
    assert_true(Xlsxrb::NumberFormatter.date_format_for?(1, styles))
    assert_true(Xlsxrb::NumberFormatter.date_only_format_for?(1, styles))
    assert_false(Xlsxrb::NumberFormatter.datetime_format_for?(1, styles))
    assert_false(Xlsxrb::NumberFormatter.time_format_for?(1, styles))

    assert_true(Xlsxrb::NumberFormatter.datetime_format_for?(3, styles))
    assert_true(Xlsxrb::NumberFormatter.time_format_for?(2, styles))
  end

  test "Ooxml::StylesParser.parse returns Elements::Styles with precomputed classifications" do
    xml = <<~XML
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts count="1">
          <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
        </numFmts>
        <cellXfs count="2">
          <xf numFmtId="0"/>
          <xf numFmtId="164"/>
        </cellXfs>
      </styleSheet>
    XML

    parsed = Xlsxrb::Ooxml::StylesParser.parse(xml)
    assert_instance_of(Xlsxrb::Elements::Styles, parsed)
    assert_equal(2, parsed[:cell_xfs].size)
    assert_false(parsed.date_format?(0))
    assert_true(parsed.date_format?(1))
    assert_equal("yyyy-mm-dd", parsed.number_format(1))
    assert_equal(:date, parsed.format_type(1))

    # Empty XML returns empty Elements::Styles
    empty_parsed = Xlsxrb::Ooxml::StylesParser.parse("")
    assert_instance_of(Xlsxrb::Elements::Styles, empty_parsed)
    assert_nil(empty_parsed.number_format(0))

    # Test format_code_for and classification_for aliases
    assert_equal("yyyy-mm-dd", parsed.format_code_for(1))
    assert_equal(:date, parsed.classification_for(1))
  end

  test "NumberFormatter checkers accept optional num_fmt_id" do
    assert_true(Xlsxrb::NumberFormatter.date_format?(nil, 14))
    assert_true(Xlsxrb::NumberFormatter.date_only_format?(nil, 14))
    assert_false(Xlsxrb::NumberFormatter.time_format?(nil, 14))
    assert_false(Xlsxrb::NumberFormatter.datetime_format?(nil, 14))

    assert_true(Xlsxrb::NumberFormatter.date_format?(nil, 20))
    assert_true(Xlsxrb::NumberFormatter.time_format?(nil, 20))
    assert_false(Xlsxrb::NumberFormatter.date_only_format?(nil, 20))

    assert_true(Xlsxrb::NumberFormatter.date_format?(nil, 22))
    assert_true(Xlsxrb::NumberFormatter.datetime_format?(nil, 22))

    assert_false(Xlsxrb::NumberFormatter.date_format?(nil, 1))
    assert_false(Xlsxrb::NumberFormatter.time_format?(nil, 1))
  end

  test "Elements::Cell format classification methods" do
    date_cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 0, value: 44_927, format_code: "yyyy-mm-dd")
    assert_equal(:date, date_cell.format_type)
    assert_true(date_cell.date_format?)
    assert_false(date_cell.time_format?)
    assert_equal(:date, date_cell[:format_type])
    assert_true(date_cell[:date_format])
    assert_false(date_cell[:time_format])

    time_cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 1, value: 0.5, format_code: "hh:mm:ss")
    assert_equal(:time, time_cell.format_type)
    assert_true(time_cell.date_format?)
    assert_true(time_cell.time_format?)

    num_cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 2, value: 123, format_code: "#,##0.00")
    assert_equal(:number, num_cell.format_type)
    assert_false(num_cell.date_format?)
    assert_false(num_cell.time_format?)

    plain_cell = Xlsxrb::Elements::Cell.new(row_index: 0, column_index: 3, value: "hello")
    assert_nil(plain_cell.format_type)
    assert_false(plain_cell.date_format?)
    assert_false(plain_cell.time_format?)
  end
end
