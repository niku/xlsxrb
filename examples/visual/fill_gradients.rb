# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "fill_gradients.xlsx"
Xlsxrb.write(output_path) do |wb|
  wb.style("gradient") do |style|
    style.fill_gradient(type: "linear", degree: 45, stops: [
                          { position: 0, color: "FFFFFFFF" },
                          { position: 1, color: "FF4F81BD" }
                        ])
  end
  wb.sheet("Gradients") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.row(["Normal Cell", "Gradient Cell"], styles: { 1 => "gradient" })
  end
end
