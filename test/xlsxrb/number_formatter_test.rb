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
end
