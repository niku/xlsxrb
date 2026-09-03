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
end
