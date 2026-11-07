# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents an embedded image anchored to worksheet cells.
    # Coordinates (from_row, from_column, to_row, to_column) are 0-based.
    #
    # @example Access image properties
    #   image = sheet.images.first
    #   image.filename       #=> "image1.png"
    #   image.ref            #=> "B3"
    #   image.content_type   #=> "image/png"
    #   image.data           #=> "\x89PNG..."
    #
    # @api public
    class Image < Data.define(:name, :filename, :content_type, :data, :from_column, :from_row, :to_column, :to_row, :cx, :cy, :unmapped_data)
      MIME_TYPES = {
        "png" => "image/png",
        "jpg" => "image/jpeg",
        "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "svg" => "image/svg+xml",
        "tif" => "image/tiff",
        "tiff" => "image/tiff",
        "bmp" => "image/bmp",
        "webp" => "image/webp",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf"
      }.freeze

      # @param filename [String] Target filename (e.g. "image1.png").
      # @param name [String, nil] Image or shape name.
      # @param content_type [String, nil] MIME content type (e.g. "image/png").
      # @param data [String, nil] Raw binary image data.
      # @param from_column [Integer, nil] 0-based starting column index.
      # @param from_row [Integer] 0-based starting row index.
      # @param to_column [Integer, nil] 0-based ending column index (nil for 1-cell anchor).
      # @param to_row [Integer, nil] 0-based ending row index (nil for 1-cell anchor).
      # @param cx [Integer, nil] Extent width in EMUs (English Metric Units).
      # @param cy [Integer, nil] Extent height in EMUs (English Metric Units).
      # @param unmapped_data [Hash] Additional metadata from drawing XML.
      #: (filename: String, ?name: String?, ?content_type: String?, ?data: String?, ?from_column: Integer?, ?from_row: Integer, ?to_column: Integer?, ?to_row: Integer?, ?cx: Integer?, ?cy: Integer?, ?unmapped_data: Hash[untyped, untyped]) -> void
      # rubocop:disable Naming/MethodParameterName
      def initialize(filename:, name: nil, content_type: nil, data: nil, from_column: nil, from_row: 0,
                     to_column: nil, to_row: nil, cx: nil, cy: nil, unmapped_data: EMPTY_HASH)
        # rubocop:enable Naming/MethodParameterName
        from_column ||= 0
        ext = File.extname(filename).delete_prefix(".").downcase
        mime = content_type || MIME_TYPES[ext] || "application/octet-stream"
        super(
          name: name,
          filename: filename,
          content_type: mime,
          data: data&.b,
          from_column: from_column,
          from_row: from_row,
          to_column: to_column,
          to_row: to_row,
          cx: cx,
          cy: cy,
          unmapped_data: unmapped_data
        )
      end

      # Returns the top-left cell coordinate reference string (e.g. "A1").
      #
      # @return [String]
      # @api public
      #: () -> String
      def ref
        "#{Cell._column_letter(from_column)}#{from_row + 1}"
      end

      # Returns the bottom-right cell coordinate reference string (e.g. "B5"), or nil for 1-cell anchors.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def to_cell_ref
        return nil if to_column.nil? || to_row.nil?

        "#{Cell._column_letter(to_column)}#{to_row + 1}"
      end

      # Returns whether this image uses a one-cell anchor.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def one_cell_anchor?
        to_column.nil? || to_row.nil?
      end

      # Returns whether this image uses a two-cell anchor.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def two_cell_anchor?
        !one_cell_anchor?
      end

      # Returns the size of the binary image data in bytes.
      #
      # @return [Integer]
      # @api public
      #: () -> Integer
      def bytesize
        data ? data.bytesize : 0
      end

      # Writes the image data directly to a file on disk.
      #
      # @param path [String] Target file path.
      # @return [Integer] Number of bytes written.
      # @api public
      #: (String path) -> Integer
      def write_to(path)
        raise Error, "No image binary data available to write" unless data

        File.binwrite(path, data)
      end
    end
  end
end
