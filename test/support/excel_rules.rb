# frozen_string_literal: true

require "rexml/document"
require "stringio"

# Checks an XLSX package for rules that Excel enforces but the ECMA-376 schemas do not express,
# such as index ranges, relationships between parts, and pivot table consistency. A file that
# breaks one of them can pass schema validation and still make Excel repair it or crash.
module ExcelRules
  Violation = Data.define(:rule, :part, :message) do
    def to_s = "#{rule} #{part}: #{message}"
  end

  BUILTIN_TABLE_STYLE = /\ATableStyle(Light([1-9]|1\d|2[01])|Medium([1-9]|1\d|2[0-8])|Dark([1-9]|1[01]))\z/
  BUILTIN_PIVOT_STYLE = /\APivotStyle(Light|Medium|Dark)([1-9]|1\d|2[0-8])\z/

  # Package parts keyed by name. OPC compares part names case-insensitively, so lookups ignore case.
  class PartNames
    def initialize(parts)
      @parts = parts
      @by_lower = parts.keys.to_h { |name| [name.downcase, name] }
    end

    def keys = @parts.keys
    def key?(name) = @by_lower.key?(name.downcase)
    def [](name) = @parts[@by_lower[name.downcase]]
    def fetch(name) = self[name] || raise(KeyError, "part not found: #{name}")
  end

  module_function

  # Returns the violations found in the XLSX bytes +data+.
  def check(data)
    parts = PartNames.new(Xlsxrb::Ooxml::ZipReader.new(StringIO.new(data)).read_all)
    docs = Hash.new { |cache, name| cache[name] = REXML::Document.new(parts.fetch(name)) }
    [
      duplicate_entries(data),
      relationship_targets(parts, docs),
      content_types(parts, docs),
      style_references(parts, docs),
      pivot_tables(parts, docs),
      style_names(parts, docs),
      comment_notes(parts, docs)
    ].flatten
  end

  # Each part is stored once: Excel rejects a package with two entries of the same name.
  def duplicate_entries(data)
    names = central_directory_names(data)
    names.tally.select { |_, count| count > 1 }.map { |name, count| Violation.new(:duplicate_entry, name, "stored #{count} times") }
  end

  # Every internal relationship points to a part that exists.
  def relationship_targets(parts, docs)
    parts.keys.grep(%r{(\A|/)_rels/[^/]*\.rels\z}).flat_map do |rels|
      source_dir = File.dirname(rels, 2)
      elements(docs[rels], "Relationship").filter_map do |rel|
        next if rel.attributes["TargetMode"] == "External"

        target = resolve(source_dir, rel.attributes["Target"])
        Violation.new(:relationship_target, rels, "#{rel.attributes["Id"]} points to missing #{target}") unless parts.key?(target)
      end
    end
  end

  # Every part has a content type, by Override or by the Default for its extension.
  def content_types(parts, docs)
    return [Violation.new(:content_type, "[Content_Types].xml", "is missing")] unless parts.key?("[Content_Types].xml")

    types = docs["[Content_Types].xml"]
    defaults = elements(types, "Default").map { |d| d.attributes["Extension"].downcase }
    overrides = elements(types, "Override").map { |o| o.attributes["PartName"].delete_prefix("/").downcase }
    parts.keys.reject { |name| name.casecmp?("[Content_Types].xml") || overrides.include?(name.downcase) || defaults.include?(name[%r{\.([^./]+)\z}, 1].to_s.downcase) }
         .map { |name| Violation.new(:content_type, name, "has no content type") }
  end

  # Style indexes stay within the lists they point into.
  def style_references(parts, docs)
    return [] unless parts.key?("xl/styles.xml")

    styles = docs["xl/styles.xml"]
    count = ->(list, item) { elements(styles, list).first&.then { |l| l.elements.to_a.count { |e| e.name == item } } || 0 }
    fonts = count.call("fonts", "font")
    fills = count.call("fills", "fill")
    borders = count.call("borders", "border")
    style_xfs = count.call("cellStyleXfs", "xf")
    cell_xfs = count.call("cellXfs", "xf")
    dxfs = count.call("dxfs", "dxf")
    custom_formats = elements(styles, "numFmt").map { |f| f.attributes["numFmtId"].to_i }
    violations = []
    check_index = lambda do |part, what, value, limit|
      violations << Violation.new(:style_index, part, "#{what} #{value} is outside 0...#{limit}") if value && value.to_i >= limit
    end
    %w[cellXfs cellStyleXfs].each do |list|
      list_node = elements(styles, list).first
      xfs = list_node ? list_node.elements.to_a : []
      xfs.each_with_index do |xf, i|
        at = "#{list}[#{i}]"
        check_index.call("xl/styles.xml", "#{at} fontId", xf.attributes["fontId"], fonts)
        check_index.call("xl/styles.xml", "#{at} fillId", xf.attributes["fillId"], fills)
        check_index.call("xl/styles.xml", "#{at} borderId", xf.attributes["borderId"], borders)
        check_index.call("xl/styles.xml", "#{at} xfId", xf.attributes["xfId"], style_xfs) if list == "cellXfs"
        format_id = xf.attributes["numFmtId"].to_i
        violations << Violation.new(:style_index, "xl/styles.xml", "#{at} numFmtId #{format_id} is not defined") if format_id >= 164 && !custom_formats.include?(format_id)
      end
    end
    elements(styles, "cellStyle").each { |cs| check_index.call("xl/styles.xml", "cellStyle #{cs.attributes["name"].inspect} xfId", cs.attributes["xfId"], style_xfs) }
    parts.keys.grep(%r{\Axl/worksheets/[^/]+\.xml\z}).each do |sheet|
      xml = parts[sheet]
      xml.scan(/<c [^>]*?\bs="(\d+)"/) { |(s)| check_index.call(sheet, "cell style", s, cell_xfs) }
      xml.scan(/<(?:row|col) [^>]*?\b(?:s|style)="(\d+)"/) { |(s)| check_index.call(sheet, "row or column style", s, cell_xfs) }
      xml.scan(/\bdxfId="(\d+)"/) { |(d)| check_index.call(sheet, "dxfId", d, dxfs) }
    end
    violations.uniq
  end

  # Pivot table field lists agree with their cache and with Excel's subtotal rules.
  def pivot_tables(parts, docs)
    parts.keys.grep(%r{\Axl/pivotTables/[^/]+\.xml\z}).flat_map do |part|
      table = docs[part]
      cache_fields = pivot_cache_fields(part, parts, docs)
      fields = elements(table, "pivotField")
      violations = []
      pivot_fields = elements(table, "pivotFields").first
      declared = pivot_fields && pivot_fields.attributes["count"]&.to_i
      violations << Violation.new(:pivot_fields, part, "pivotFields count #{declared} but #{fields.size} pivotField") if declared && declared != fields.size
      violations << Violation.new(:pivot_fields, part, "#{fields.size} pivotField but #{cache_fields.size} cacheField") if cache_fields && cache_fields.size != fields.size
      axis_fields = %w[rowFields colFields].flat_map { |list| elements(table, list).flat_map { |l| elements(l, "field").map { |f| f.attributes["x"].to_i } } }.reject(&:negative?)
      axis_fields.each do |index|
        field = fields[index]
        violations << Violation.new(:pivot_axis_items, part, "row or column field #{index} has no <items>") if field && elements(field, "items").empty?
      end
      fields.each_with_index do |field, index|
        items_list = elements(field, "items").first or next
        items = elements(items_list, "item")
        declared_items = items_list.attributes["count"]&.to_i
        violations << Violation.new(:pivot_items, part, "field #{index} items count #{declared_items} but #{items.size} item") if declared_items && declared_items != items.size
        violations << Violation.new(:pivot_default_subtotal, part, "field #{index} has defaultSubtotal=\"0\" and a default item") if field.attributes["defaultSubtotal"] == "0" && items.any? { |i| i.attributes["t"] == "default" }
        shared = cache_fields && cache_fields[index]
        next unless shared

        too_far = items.map { |i| i.attributes["x"] }.compact.map(&:to_i).select { |x| x >= shared }
        violations << Violation.new(:pivot_items, part, "field #{index} items #{too_far.inspect} are outside its #{shared} shared items") unless too_far.empty?
      end
      violations
    end
  end

  # Table and pivot table style names are built in or defined in styles.xml.
  def style_names(parts, docs)
    custom = parts.key?("xl/styles.xml") ? elements(docs["xl/styles.xml"], "tableStyle").map { |s| s.attributes["name"] } : []
    checks = { %r{\Axl/tables/[^/]+\.xml\z} => ["tableStyleInfo", BUILTIN_TABLE_STYLE], %r{\Axl/pivotTables/[^/]+\.xml\z} => ["pivotTableStyleInfo", BUILTIN_PIVOT_STYLE] }
    checks.flat_map do |pattern, (element, builtin)|
      parts.keys.grep(pattern).flat_map do |part|
        elements(docs[part], element).filter_map do |info|
          name = info.attributes["name"]
          next if name.nil? || name.match?(builtin) || custom.include?(name)

          Violation.new(:style_name, part, "#{element} name #{name.inspect} is neither built in nor defined")
        end
      end
    end
  end

  # Each comment has a note shape in the sheet's legacy VML drawing at the same cell.
  def comment_notes(parts, docs)
    parts.keys.grep(%r{\Axl/worksheets/_rels/[^/]+\.rels\z}).flat_map do |rels|
      targets = elements(docs[rels], "Relationship").to_h { |r| [r.attributes["Type"].split("/").last, resolve("xl/worksheets", r.attributes["Target"])] }
      comments = targets["comments"]
      vml = targets["vmlDrawing"]
      next [] unless comments && parts.key?(comments)
      next [Violation.new(:comment_notes, comments, "comments without a VML drawing")] unless vml && parts.key?(vml)

      notes = parts[vml].split("</v:shape>").filter_map do |shape|
        next unless shape.include?('ObjectType="Note"')

        row = shape[%r{<x:Row>\s*(\d+)\s*</x:Row>}, 1]
        col = shape[%r{<x:Column>\s*(\d+)\s*</x:Column>}, 1]
        "#{Xlsxrb.column_letter(col.to_i)}#{row.to_i + 1}" if row && col
      end
      elements(docs[comments], "comment").map { |c| c.attributes["ref"] }.reject { |ref| notes.include?(ref) }
                                         .map { |ref| Violation.new(:comment_notes, vml, "no note for comment #{ref}") }
    end
  end

  def pivot_cache_fields(pivot_part, parts, docs)
    rels = "#{File.dirname(pivot_part)}/_rels/#{File.basename(pivot_part)}.rels"
    return unless parts.key?(rels)

    target = elements(docs[rels], "Relationship").find { |r| r.attributes["Type"].end_with?("/pivotCacheDefinition") }
    cache = target && resolve(File.dirname(pivot_part), target.attributes["Target"])
    return unless cache && parts.key?(cache)

    elements(docs[cache], "cacheField").map { |f| elements(f, "sharedItems").first&.elements&.size.to_i }
  end

  def elements(node, name)
    REXML::XPath.match(node, ".//*[local-name()='#{name}']")
  end

  def resolve(source_dir, target)
    return target.delete_prefix("/") if target.start_with?("/")

    parts = (source_dir.empty? || source_dir == "." ? [] : source_dir.split("/"))
    target.split("/").each { |segment| segment == ".." ? parts.pop : parts << segment }
    parts.join("/")
  end

  # Entry names from the central directory, duplicates included (ZipReader keeps the last one).
  def central_directory_names(data)
    data = data.b
    eocd = data.rindex("PK\x05\x06".b) or return []
    count, size, offset = data.byteslice(eocd + 10, 10).unpack("vVV")
    return [] if [count, size, offset].intersect?([0xFFFF, 0xFFFFFFFF])

    pos = offset
    Array.new(count) do
      name_length, extra_length, comment_length = data.byteslice(pos + 28, 6).unpack("vvv")
      name = data.byteslice(pos + 46, name_length).force_encoding("UTF-8")
      pos += 46 + name_length + extra_length + comment_length
      name
    end
  end
end
