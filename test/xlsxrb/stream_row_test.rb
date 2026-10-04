# frozen_string_literal: true

require "test_helper"

class StreamRowTest < Test::Unit::TestCase
  cover Xlsxrb::StreamRow

  test "stream_row parses sparse cells and fills array with nil" do
    xml = %(<c r="A1"><v>42</v></c><c r="D1" t="s"><v>0</v></c>).b
    row = Xlsxrb::StreamRow.new(
      index: 0,
      xml_bytes: xml,
      from: 0,
      to: xml.bytesize,
      shared_strings: ["world"],
      height: 22.5,
      hidden: true,
      custom_height: true,
      outline_level: 1
    )

    # Basic attributes
    assert_equal(0, row.index)
    assert_equal(22.5, row.height)
    assert_equal(true, row.hidden)
    assert_equal(true, row.custom_height)
    assert_equal(1, row.outline_level)
    assert_equal(true, row.valid?)
    assert_equal([], row.errors)
    assert_equal({}, row.unmapped_data)
    assert_match(/index=0/, row.inspect)

    # Symbol bracket access
    assert_equal(row.cells, row[:cells])
    assert_equal(0, row[:index])
    assert_equal(22.5, row[:height])
    assert_equal(true, row[:hidden])
    assert_equal(true, row[:custom_height])
    assert_equal(1, row[:outline_level])
    assert_equal({ height: 22.5, hidden: true, custom_height: true, outline_level: 1 }, row[:attrs])

    # Integer bracket access
    assert_equal(42, row[0].value)
    assert_equal("world", row[1].value) # second parsed cell is col 3

    # cell_at 0-based col index
    assert_equal(42, row.cell_at(0).value)
    assert_nil(row.cell_at(1))
    assert_nil(row.cell_at(2))
    assert_equal("world", row.cell_at(3).value)

    # to_a and values sparse array
    assert_equal([42, nil, nil, "world"], row.to_a)
    assert_equal([42, nil, nil, "world"], row.values)

    # Enumerable each and each_cell
    assert_equal(2, row.count)
    cells_collected = []
    row.each_cell { |c| cells_collected << c.value }
    assert_equal([42, "world"], cells_collected)
  end

  test "stream_row with empty cells returns empty array" do
    empty_xml = "".b
    row = Xlsxrb::StreamRow.fast_create(5, empty_xml, 0, 0, [])

    assert_equal(5, row.index)
    assert_equal([], row.cells)
    assert_equal([], row.to_a)
    assert_equal([], row.values)
    assert_nil(row.cell_at(0))
  end

  test "stream_row unmapped_data returns hash with style_index when present and EMPTY_HASH when nil" do
    row_without_style = Xlsxrb::StreamRow.new(
      index: 0,
      xml_bytes: "".b,
      from: 0,
      to: 0,
      shared_strings: [],
      style_index: nil
    )
    assert_same(Xlsxrb::Elements::EMPTY_HASH, row_without_style.unmapped_data)
    assert_equal({}, row_without_style.unmapped_data)

    row_with_style = Xlsxrb::StreamRow.new(
      index: 0,
      xml_bytes: "".b,
      from: 0,
      to: 0,
      shared_strings: [],
      style_index: 3
    )
    assert_equal({ style_index: 3 }, row_with_style.unmapped_data)
    assert_equal(3, row_with_style.unmapped_data[:style_index])
    assert_equal(3, row_with_style.style_index)
    assert_equal(3, row_with_style[:style_index])
    assert_equal({ height: nil, hidden: false, custom_height: false, outline_level: nil, style_index: 3 }, row_with_style[:attrs])
  end

  test "stream_row parses shared formula cells and retains values" do
    xml = %(<c r="A1" s="2"><f t="shared" ref="A1:A5" si="0">B1+C1</f><v>10</v></c><c r="B1" s="2"><f t="shared" si="0"/><v>20</v></c><c r="C1" s="2"><f t="shared" si="0" /><v>30</v></c>).b
    row = Xlsxrb::StreamRow.new(
      index: 0,
      xml_bytes: xml,
      from: 0,
      to: xml.bytesize,
      shared_strings: []
    )

    assert_equal(3, row.cells.size)
    assert_equal("A1", row.cells[0].ref)
    assert_equal(10, row.cells[0].value)
    assert_equal("B1+C1", row.cells[0].formula_expression)

    assert_equal("B1", row.cells[1].ref)
    assert_equal(20, row.cells[1].value)
    assert_nil(row.cells[1].formula_expression)

    assert_equal("C1", row.cells[2].ref)
    assert_equal(30, row.cells[2].value)
    assert_nil(row.cells[2].formula_expression)

    assert_equal([10, 20, 30], row.values)
  end

  test "stream_row date1904? reflects date1904 flag and bracket access" do
    row1900 = Xlsxrb::StreamRow.fast_create(0, "".b, 0, 0, [], "", nil, false, false, nil, nil, nil, false)
    assert_equal(false, row1900.date1904?)
    assert_equal(false, row1900[:date1904])

    row1904 = Xlsxrb::StreamRow.fast_create(0, "".b, 0, 0, [], "", nil, false, false, nil, nil, nil, true)
    assert_equal(true, row1904.date1904?)
    assert_equal(true, row1904[:date1904])
  end

  test "stream_row values parses XML directly without creating Cell objects when cells is nil" do
    xml = %(<c r="A1"><v>42</v></c><c r="B1" t="s"><v>0</v></c><c r="D1"><v>100</v></c>).b
    row = Xlsxrb::StreamRow.new(
      index: 0,
      xml_bytes: xml,
      from: 0,
      to: xml.bytesize,
      shared_strings: ["test"]
    )
    assert_nil(row.instance_variable_get(:@cells))
    assert_equal([42, "test", nil, 100], row.values)
    assert_nil(row.instance_variable_get(:@cells))
  end

  test "stream_row values uses cached cells when @cells is present" do
    row = Xlsxrb::StreamRow.fast_create(0, "".b, 0, 0, [])
    cell = Xlsxrb::Elements::Cell.fast_create(0, 0, "from_cached_cells", nil, nil, nil, nil, false)
    row.instance_variable_set(:@cells, [cell])
    assert_equal(["from_cached_cells"], row.values)
  end

  test "stream_row values default type_cast is false and leaves date serial as numeric" do
    xml = %(<c r="A1" s="1"><v>44927</v></c>).b
    styles = Xlsxrb::Elements::Styles.new({
                                            cell_xfs: [{ num_fmt_id: 0 }, { num_fmt_id: 14 }]
                                          })
    row = Xlsxrb::StreamRow.new(
      index: 0,
      xml_bytes: xml,
      from: 0,
      to: xml.bytesize,
      shared_strings: [],
      styles: styles
    )
    assert_equal([44_927], row.values)
    assert_equal([44_927], row.values(type_cast: false))
    assert_equal([Date.new(2023, 1, 1)], row.values(type_cast: true))
  end

  test "stream_row values cached cells supports type_cast true and false with date1904" do
    xml = %(<c r="A1" s="1"><v>44927</v></c><c r="B1" s="2"><v>44927.5</v></c><c r="C1" s="3"><v>0.5</v></c><c r="D1" s="1" t="str"><v>not_num</v></c>).b
    styles = Xlsxrb::Elements::Styles.new({
                                            cell_xfs: [
                                              { num_fmt_id: 0 },
                                              { num_fmt_id: 14 },
                                              { num_fmt_id: 22 },
                                              { num_fmt_id: 20 }
                                            ]
                                          })
    row1900 = Xlsxrb::StreamRow.new(
      index: 0,
      xml_bytes: xml,
      from: 0,
      to: xml.bytesize,
      shared_strings: [],
      styles: styles,
      date1904: false
    )
    assert_equal(4, row1900.cells.size)
    assert_not_nil(row1900.instance_variable_get(:@cells))

    assert_equal([44_927, 44_927.5, 0.5, "not_num"], row1900.values(type_cast: false))
    values1900 = row1900.values(type_cast: true)
    assert_equal(Date.new(2023, 1, 1), values1900[0])
    assert_equal(Time.utc(2023, 1, 1, 12, 0, 0), values1900[1])
    assert_equal(Time.utc(1899, 12, 31, 12, 0, 0), values1900[2])
    assert_equal("not_num", values1900[3])

    xml1904 = %(<c r="A1" s="1"><v>43465</v></c>).b
    row1904 = Xlsxrb::StreamRow.new(
      index: 0,
      xml_bytes: xml1904,
      from: 0,
      to: xml1904.bytesize,
      shared_strings: [],
      styles: styles,
      date1904: true
    )
    assert_equal([Date.new(2023, 1, 1)], row1904.values(type_cast: true))
  end
end
