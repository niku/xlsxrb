# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Ooxml
    # Pure-Ruby Compound File Binary (CFB / OLE Structured Storage) implementation for [MS-CFB] / [MS-OFFCRYPTO].
    module Cfb
      MAGIC = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1".b.freeze

      FREESECT   = 0xFFFFFFFF
      ENDOFCHAIN = 0xFFFFFFFE
      FATSECT    = 0xFFFFFFFD
      DIFATSECT  = 0xFFFFFFFC
      NOSTREAM   = 0xFFFFFFFF

      OBJ_UNKNOWN = 0x00
      OBJ_STORAGE = 0x01
      OBJ_STREAM  = 0x02
      OBJ_ROOT    = 0x05

      MINI_STREAM_CUTOFF = 4096
      SECTOR_SIZE = 512
      MINI_SECTOR_SIZE = 64

      # Represents a directory entry in a Compound File.
      class DirEntry
        attr_accessor :name, :type, :color, :left_sibling_id, :right_sibling_id, :child_id, :clsid, :state_flags, :created_time, :modified_time, :start_sector, :size, :entry_id

        def initialize(name: "", type: OBJ_UNKNOWN, start_sector: ENDOFCHAIN, size: 0)
          @name = name
          @type = type
          @color = 0 # Red / Black (0 is acceptable across OLE implementations)
          @left_sibling_id = NOSTREAM
          @right_sibling_id = NOSTREAM
          @child_id = NOSTREAM
          @clsid = "\x00".b * 16
          @state_flags = 0
          @created_time = 0
          @modified_time = 0
          @start_sector = start_sector
          @size = size
          @entry_id = 0
        end

        def stream?
          @type == OBJ_STREAM
        end

        def root?
          @type == OBJ_ROOT
        end

        def storage?
          @type == OBJ_STORAGE
        end
      end

      # Reads streams from a Compound File Binary buffer.
      class Reader
        attr_reader :entries

        def self.cfb?(data)
          return false if data.nil? || data.bytesize < 8

          data[0, 8] == MAGIC
        end

        def initialize(data)
          @data = data.b
          raise Xlsxrb::Error, "Invalid CFB file header signature" unless Reader.cfb?(@data)

          parse_header
          build_fat
          parse_directory
          load_mini_stream
        end

        def stream_names
          @entries.select(&:stream?).map(&:name)
        end

        def read_stream(name)
          entry = @entries.find { |e| e.name.casecmp?(name) && e.stream? }
          return nil unless entry

          if entry.size < @mini_cutoff && @mini_stream && !@minifat.empty?
            read_mini_stream_data(entry.start_sector, entry.size)
          else
            read_regular_stream_data(entry.start_sector, entry.size)
          end
        end

        private

        def parse_header
          raise Xlsxrb::Error, "CFB file is too small" if @data.bytesize < 512

          _, major_ver = @data[0x18, 4].unpack("v2")
          @major_version = major_ver
          @sector_shift = @data[0x1E, 2].unpack1("v")
          @mini_sector_shift = @data[0x20, 2].unpack1("v")
          @sector_size = 1 << @sector_shift
          @mini_sector_size = 1 << @mini_sector_shift

          @num_dir_sectors = @data[0x28, 4].unpack1("V")
          @num_fat_sectors = @data[0x2C, 4].unpack1("V")
          @first_dir_sector = @data[0x30, 4].unpack1("V")
          @mini_cutoff = @data[0x38, 4].unpack1("V")
          @first_minifat_sector = @data[0x3C, 4].unpack1("V")
          @num_minifat_sectors = @data[0x40, 4].unpack1("V")
          @first_difat_sector = @data[0x44, 4].unpack1("V")
          @num_difat_sectors = @data[0x48, 4].unpack1("V")

          @header_difat = @data[0x4C, 436].unpack("V109").reject { |s| [FREESECT, ENDOFCHAIN].include?(s) }
        end

        def sector_offset(sector_id)
          512 + (sector_id * @sector_size)
        end

        def read_sector(sector_id)
          offset = sector_offset(sector_id)
          @data[offset, @sector_size] || "".b
        end

        def build_fat
          load_difat_and_fat
        end

        def load_difat_and_fat
          @difat_sectors = @header_difat.dup
          if @first_difat_sector != ENDOFCHAIN && @first_difat_sector != FREESECT
            curr = @first_difat_sector
            visited = {}
            while curr != ENDOFCHAIN && curr != FREESECT
              break if visited[curr]

              visited[curr] = true
              sec_data = read_sector(curr)
              entries_per_sec = (@sector_size / 4) - 1
              entries = sec_data[0, entries_per_sec * 4].unpack("V*")
              @difat_sectors.concat(entries)
              curr = sec_data[entries_per_sec * 4, 4].unpack1("V")
            end
          end

          # Read all FAT sectors
          @fat = []
          @difat_sectors.each do |fat_sec_id|
            break if [ENDOFCHAIN, FREESECT].include?(fat_sec_id)

            sec_data = read_sector(fat_sec_id)
            @fat.concat(sec_data.unpack("V*"))
          end
        end

        def parse_directory
          dir_data = +""
          curr = @first_dir_sector
          visited = {}
          while curr != ENDOFCHAIN && curr != FREESECT && curr < @fat.size
            break if visited[curr]

            visited[curr] = true
            dir_data << read_sector(curr)
            curr = @fat[curr]
          end

          @entries = []
          num_entries = dir_data.bytesize / 128
          num_entries.times do |i|
            entry_bytes = dir_data[i * 128, 128]
            next if entry_bytes.nil? || entry_bytes.bytesize < 128

            name_bytes = entry_bytes[0, 64]
            name_len = entry_bytes[0x40, 2].unpack1("v")
            type = entry_bytes[0x42, 1].ord
            next if type == OBJ_UNKNOWN

            name_str = if name_len > 2
                         name_bytes[0, name_len - 2].force_encoding("UTF-16LE").encode("UTF-8", invalid: :replace, undef: :replace)
                       else
                         ""
                       end

            entry = DirEntry.new(name: name_str, type: type)
            entry.entry_id = i
            entry.color = entry_bytes[0x43, 1].ord
            entry.left_sibling_id = entry_bytes[0x44, 4].unpack1("V")
            entry.right_sibling_id = entry_bytes[0x48, 4].unpack1("V")
            entry.child_id = entry_bytes[0x4C, 4].unpack1("V")
            entry.clsid = entry_bytes[0x50, 16]
            entry.state_flags = entry_bytes[0x60, 4].unpack1("V")
            entry.start_sector = entry_bytes[0x74, 4].unpack1("V")
            entry.size = entry_bytes[0x78, 8].unpack1("Q<")

            @entries << entry
          end
        end

        def load_mini_stream
          @minifat = []
          if @first_minifat_sector != ENDOFCHAIN && @first_minifat_sector != FREESECT
            curr = @first_minifat_sector
            visited = {}
            while curr != ENDOFCHAIN && curr != FREESECT && curr < @fat.size
              break if visited[curr]

              visited[curr] = true
              sec_data = read_sector(curr)
              @minifat.concat(sec_data.unpack("V*"))
              curr = @fat[curr]
            end
          end

          root_entry = @entries.find(&:root?)
          @mini_stream = if root_entry && root_entry.start_sector != ENDOFCHAIN && root_entry.size.positive?
                           read_regular_stream_data(root_entry.start_sector, root_entry.size)
                         else
                           "".b
                         end
        end

        def read_regular_stream_data(start_sector, total_size)
          return "".b if [ENDOFCHAIN, FREESECT].include?(start_sector) || total_size.zero?

          result = +""
          curr = start_sector
          visited = {}
          while curr != ENDOFCHAIN && curr != FREESECT && curr < @fat.size
            break if visited[curr]

            visited[curr] = true
            result << read_sector(curr)
            break if result.bytesize >= total_size

            curr = @fat[curr]
          end
          result[0, total_size] || "".b
        end

        def read_mini_stream_data(start_mini_sector, total_size)
          return "".b if [ENDOFCHAIN, FREESECT].include?(start_mini_sector) || total_size.zero?

          result = +""
          curr = start_mini_sector
          visited = {}
          while curr != ENDOFCHAIN && curr != FREESECT && curr < @minifat.size
            break if visited[curr]

            visited[curr] = true
            offset = curr * @mini_sector_size
            result << (@mini_stream[offset, @mini_sector_size] || "".b)
            break if result.bytesize >= total_size

            curr = @minifat[curr]
          end
          result[0, total_size] || "".b
        end
      end

      # Writes named streams into a Compound File Binary (v3, 512-byte sectors) format with Mini Stream support.
      class Writer
        def self.write(streams)
          new(streams).build
        end

        def initialize(streams)
          # streams: Hash of { String => String (bytes) }
          @streams = streams
        end

        def build
          sector_size = SECTOR_SIZE
          mini_sector_size = MINI_SECTOR_SIZE

          # Partition streams into mini streams (< 4096 bytes) and regular streams (>= 4096 bytes)
          regular_stream_entries = []

          mini_stream_bytes = +""
          minifat = []

          # Directory Entries array
          dir_entries = []
          root_entry = DirEntry.new(name: "Root Entry", type: OBJ_ROOT, start_sector: ENDOFCHAIN, size: 0)
          dir_entries << root_entry

          stream_names = @streams.keys
          stream_names.each_with_index do |name, idx|
            entry_id = idx + 1
            data = @streams[name].b
            size = data.bytesize

            if size < MINI_STREAM_CUTOFF && size.positive?
              # Allocate in Mini Stream
              start_mini_sec = minifat.size
              num_mini_sec = (size + mini_sector_size - 1) / mini_sector_size
              num_mini_sec.times do |m_idx|
                chunk = data[m_idx * mini_sector_size, mini_sector_size] || "".b
                chunk = chunk.ljust(mini_sector_size, "\x00".b) if chunk.bytesize < mini_sector_size
                mini_stream_bytes << chunk
                minifat << (m_idx == num_mini_sec - 1 ? ENDOFCHAIN : (start_mini_sec + m_idx + 1))
              end
              entry = DirEntry.new(name: name, type: OBJ_STREAM, start_sector: start_mini_sec, size: size)
            elsif size >= MINI_STREAM_CUTOFF
              # Allocate in Regular Stream (start_sector will be assigned later)
              entry = DirEntry.new(name: name, type: OBJ_STREAM, start_sector: ENDOFCHAIN, size: size)
              regular_stream_entries << [entry, data]
            else
              entry = DirEntry.new(name: name, type: OBJ_STREAM, start_sector: ENDOFCHAIN, size: 0)
            end

            entry.entry_id = entry_id
            dir_entries << entry
          end

          # Set up directory binary tree
          root_entry.child_id = @streams.empty? ? NOSTREAM : 1
          if dir_entries.size > 1
            root_child = dir_entries[1]
            (2...dir_entries.size).each do |i|
              insert_entry_to_tree(dir_entries, root_child, dir_entries[i])
            end
          end

          # Pad directory entries to multiple of 4 (128 bytes * 4 = 512 bytes = 1 sector)
          dir_entries << DirEntry.new(type: OBJ_UNKNOWN) until (dir_entries.size % 4).zero?

          # Now allocate regular sectors:
          # 1. Regular stream data sectors
          # 2. Mini stream container sectors
          # 3. Mini FAT sectors
          # 4. Directory sector
          # 5. FAT sector
          allocated_sectors = []
          sector_chains = [] # pairs of [start_sec, num_sec]

          # 1. Regular streams
          regular_stream_entries.each do |entry, data|
            start_sec = allocated_sectors.size
            num_sec = (data.bytesize + sector_size - 1) / sector_size
            num_sec.times do |s_idx|
              chunk = data[s_idx * sector_size, sector_size] || "".b
              chunk = chunk.ljust(sector_size, "\x00".b) if chunk.bytesize < sector_size
              allocated_sectors << chunk
            end
            entry.start_sector = start_sec
            sector_chains << [start_sec, num_sec]
          end

          # 2. Mini Stream container (assigned to Root Entry)
          if mini_stream_bytes.bytesize.positive?
            start_sec = allocated_sectors.size
            num_sec = (mini_stream_bytes.bytesize + sector_size - 1) / sector_size
            num_sec.times do |s_idx|
              chunk = mini_stream_bytes[s_idx * sector_size, sector_size] || "".b
              chunk = chunk.ljust(sector_size, "\x00".b) if chunk.bytesize < sector_size
              allocated_sectors << chunk
            end
            root_entry.start_sector = start_sec
            root_entry.size = mini_stream_bytes.bytesize
            sector_chains << [start_sec, num_sec]
          end

          # 3. Mini FAT sectors
          first_minifat_sec = ENDOFCHAIN
          num_minifat_sec = 0
          if minifat.size.positive?
            first_minifat_sec = allocated_sectors.size
            minifat_bytes = minifat.pack("V*")
            num_minifat_sec = (minifat_bytes.bytesize + sector_size - 1) / sector_size
            num_minifat_sec.times do |s_idx|
              chunk = minifat_bytes[s_idx * sector_size, sector_size] || "".b
              chunk = chunk.ljust(sector_size, "\xFF".b) if chunk.bytesize < sector_size
              allocated_sectors << chunk
            end
            sector_chains << [first_minifat_sec, num_minifat_sec]
          end

          # 4. Directory sector
          dir_sector_id = allocated_sectors.size
          dir_sector_bytes = +""
          dir_entries.each do |e|
            dir_sector_bytes << serialize_dir_entry(e)
          end
          allocated_sectors << dir_sector_bytes

          # 5. FAT sector
          fat_sector_id = allocated_sectors.size
          fat = Array.new(fat_sector_id + 1, FREESECT)

          # Populate regular sector chains in FAT
          sector_chains.each do |start_sec, num_sec|
            num_sec.times do |s_idx|
              fat[start_sec + s_idx] = s_idx == num_sec - 1 ? ENDOFCHAIN : (start_sec + s_idx + 1)
            end
          end

          fat[dir_sector_id] = ENDOFCHAIN
          fat[fat_sector_id] = FATSECT

          fat_bytes = fat.pack("V*").ljust(sector_size, "\xFF".b)
          allocated_sectors << fat_bytes

          # Build Header (512 bytes)
          header = +""
          header << MAGIC # 0x00 (8 bytes)
          header << ("\x00".b * 16) # 0x08 CLSID
          header << [0x003B, 0x0003].pack("v2") # 0x18 Minor version (0x3B), Major version (3)
          header << [0xFFFE].pack("v") # 0x1C Byte order (Little Endian)
          header << [9].pack("v") # 0x1E Sector shift (512 bytes)
          header << [6].pack("v") # 0x20 Mini sector shift (64 bytes)
          header << ("\x00".b * 6) # 0x22 Reserved
          header << [0].pack("V") # 0x28 Number of Directory sectors (0 for v3)
          header << [1].pack("V") # 0x2C Number of FAT sectors (1)
          header << [dir_sector_id].pack("V") # 0x30 First Directory sector
          header << [0].pack("V") # 0x34 Transaction signature
          header << [MINI_STREAM_CUTOFF].pack("V") # 0x38 Mini stream cutoff size (4096)
          header << [first_minifat_sec].pack("V") # 0x3C First Mini FAT sector
          header << [num_minifat_sec].pack("V") # 0x40 Number of Mini FAT sectors
          header << [ENDOFCHAIN].pack("V") # 0x44 First DIFAT sector
          header << [0].pack("V") # 0x48 Number of DIFAT sectors

          # DIFAT array (109 entries * 4 bytes = 436 bytes)
          difat = Array.new(109, FREESECT)
          difat[0] = fat_sector_id
          header << difat.pack("V109")

          raise "Header size mismatch" unless header.bytesize == 512

          # Combine Header + All Sectors
          output = +""
          output << header
          allocated_sectors.each { |sec| output << sec }
          output
        end

        private

        def serialize_dir_entry(entry)
          return "\x00".b * 128 if entry.type == OBJ_UNKNOWN

          buf = +""
          # Name in UTF-16LE with null terminator
          name_utf16 = "#{entry.name}\u0000".encode("UTF-16LE")
          name_bytes = name_utf16.b[0, 64].ljust(64, "\x00".b)
          buf << name_bytes
          buf << [name_utf16.bytesize].pack("v") # 0x40 Name length
          buf << [entry.type].pack("C")          # 0x42 Object type
          buf << [entry.color].pack("C")         # 0x43 Color (0 = Red/Black)
          buf << [entry.left_sibling_id].pack("V")  # 0x44 Left sibling
          buf << [entry.right_sibling_id].pack("V") # 0x48 Right sibling
          buf << [entry.child_id].pack("V")         # 0x4C Child ID
          buf << (entry.clsid || ("\x00".b * 16))   # 0x50 CLSID
          buf << [entry.state_flags].pack("V")      # 0x60 State flags
          buf << [entry.created_time].pack("Q<")    # 0x64 Created time
          buf << [entry.modified_time].pack("Q<")   # 0x6C Modified time
          buf << [entry.start_sector].pack("V")     # 0x74 Starting sector
          buf << [entry.size].pack("Q<")            # 0x78 Stream size (uint64)

          buf.ljust(128, "\x00".b)
        end

        def insert_entry_to_tree(entries, root_node, new_node)
          cmp = compare_entry_names(new_node.name, root_node.name)
          if cmp.negative?
            if root_node.left_sibling_id == NOSTREAM
              root_node.left_sibling_id = new_node.entry_id
            else
              insert_entry_to_tree(entries, entries[root_node.left_sibling_id], new_node)
            end
          elsif root_node.right_sibling_id == NOSTREAM
            root_node.right_sibling_id = new_node.entry_id
          else
            insert_entry_to_tree(entries, entries[root_node.right_sibling_id], new_node)
          end
        end

        def compare_entry_names(a_name, b_name)
          # [MS-CFB] Section 2.6.1: Length comparison first, then uppercase UTF-16 code point comparison
          return -1 if a_name.length < b_name.length
          return 1 if a_name.length > b_name.length

          a_name.upcase <=> b_name.upcase
        end
      end
    end
  end
end
