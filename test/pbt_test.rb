# frozen_string_literal: true

require "test_helper"
require "pbt"
require "tempfile"

class PbtTest < Test::Unit::TestCase
  # Valid OOXML sheet names shouldn't contain certain characters like \/?*[]:
  # They also should be 1-31 chars long.
  def sheet_name_generator
    Pbt.alphanumeric_string(min: 1, max: 31)
  end

  # Cell values can be numbers, boolean, strings (excluding invalid XML chars), or nil
  def cell_value_generator
    Pbt.tuple(
      Pbt.choose(0..3),
      Pbt.integer,
      Pbt.boolean,
      Pbt.printable_ascii_string(max: 20)
    ).map(
      ->(t) {
        case t[0]
        when 0 then t[1]
        when 1 then t[2]
        when 2 then nil
        when 3 then t[3]
        end
      },
      ->(v) {
        case v
        when Integer then [0, v, false, ""]
        when TrueClass, FalseClass then [1, 0, v, ""]
        when NilClass then [2, 0, false, ""]
        when String then [3, 0, false, v]
        else [2, 0, false, ""]
        end
      }
    )
  end

  test "property-based test: build and read consistency" do
    Pbt.assert(num_runs: 50) do
      Pbt.property(
        Pbt.array(sheet_name_generator, min: 1, max: 3).filter { |a| a.size == a.uniq.size && !a.empty? },
        Pbt.array(Pbt.array(cell_value_generator, max: 5), max: 5)
      ) do |sheet_names, rows_data|
        workbook = Xlsxrb.build do |w|
          sheet_names.each do |sname|
            w.sheet(sname) do |s|
              rows_data.each do |row_data|
                s.row(row_data)
              end
            end
          end
        end

        tmp = Tempfile.new(["pbt_test", ".xlsx"])
        begin
          Xlsxrb.write(tmp.path, workbook)
          
          read_wb = Xlsxrb.read(tmp.path)
          
          assert_equal sheet_names.size, read_wb.sheets.size
          
          sheet_names.each_with_index do |sname, sheet_idx|
            sheet = read_wb.sheets[sheet_idx]
            assert_equal sname, sheet.name
            
            rows_data.each_with_index do |expected_row, row_idx|
              actual_row = sheet.rows.find { |r| r.index == row_idx }
              
              if expected_row.all?(&:nil?)
                next
              end
              
              expected_row.each_with_index do |expected_val, col_idx|
                actual_cell = actual_row&.cells&.find { |c| c.column_index == col_idx }
                actual_val = actual_cell&.value
                
                # Check loosely since xlsx reader might normalize empty strings to nil, or numbers differently
                expected_str = expected_val.to_s
                actual_str = actual_val.to_s
                
                assert_equal expected_str, actual_str
              end
            end
          end
        ensure
          tmp.close!
        end
      end
    end
  end
end
