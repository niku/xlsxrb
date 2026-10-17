# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "align_text_rotation.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("rot_45") { |s| s.text_rotation(45) }
  wb.style("rot_90") { |s| s.text_rotation(90) }
  wb.sheet("Rotation") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.print_options(:grid_lines, true)
    s.row(["Rotated 45", "Rotated 90"], styles: { 0 => "rot_45", 1 => "rot_90" }, height: 50)
  end
end
