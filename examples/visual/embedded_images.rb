# frozen_string_literal: true

require "xlsxrb"
output_path = ARGV[0] || "embedded_images.xlsx"

require "zlib"

def make_png(width, height, red, green, blue)
  png = "\x89PNG\r\n\x1a\n".dup.force_encoding("BINARY")
  ihdr_data = [width, height, 8, 2, 0, 0, 0].pack("N2C5")
  ihdr_chunk = "IHDR".dup.force_encoding("BINARY") + ihdr_data
  ihdr_crc = Zlib.crc32(ihdr_chunk)
  png << [ihdr_data.bytesize].pack("N") << ihdr_chunk << [ihdr_crc].pack("N")
  raw_data = Array.new(height) { "\x00".dup.force_encoding("BINARY") + ([red, green, blue].pack("C3") * width) }.join
  compressed = Zlib.deflate(raw_data)
  idat_chunk = "IDAT".dup.force_encoding("BINARY") + compressed
  idat_crc = Zlib.crc32(idat_chunk)
  png << [compressed.bytesize].pack("N") << idat_chunk << [idat_crc].pack("N")
  iend_chunk = "IEND".dup.force_encoding("BINARY")
  iend_crc = Zlib.crc32(iend_chunk)
  png << [0].pack("N") << iend_chunk << [iend_crc].pack("N")
  png
end

dummy_png = make_png(100, 100, 255, 0, 0)

Xlsxrb.generate(output_path) do |w|
  w.style("center") { |st| st.align_horizontal(:center) }
  w.sheet("Images") do |s|
    s.row(["Logo Target cell:", "", "", "Boundary"], styles: %w[left center center center])
    s.row(["", "", "", ""])
    s.row(["", "", "", ""])
    s.row(["", "", "", ""])
    s.row(["", "", "", "Boundary End"], styles: %w[left center center center])
    s.add_image(dummy_png, ext: "png", from_col: 1, from_row: 1, to_col: 3, to_row: 5)
    s.column(0, width: 20)
    s.column(1, width: 15)
    s.column(2, width: 15)
    s.column(3, width: 15)
  end
end

# 2. Read the generated sheet and print cell values + parsed embedded images details
puts "=== Read Validation ==="
reader = Xlsxrb::Ooxml::Reader.new(output_path)
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.first(5).each do |row|
  row_cells = row.cells.map { |c| "#{c.ref}: #{c.value.inspect}" }
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
images = reader.images(sheet: sheet.name)
images.each_with_index do |img, idx|
  zip_path = File.expand_path(img[:target], "/xl/drawings").sub(%r{^/}, "")
  img_data = reader.send(:extract_zip_entry, zip_path)
  is_png = img_data && img_data[0..3] == "\x89PNG".b
  puts "Image ##{idx + 1}: name='#{img[:name]}', target='#{img[:target]}' -> ZIP path='#{zip_path}', size=#{img_data&.bytesize} bytes, valid_png=#{is_png}, range=Col #{img[:from_col]} Row #{img[:from_row]} to Col #{img[:to_col]} Row #{img[:to_row]}"
end
