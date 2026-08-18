# frozen_string_literal: true
# rubocop:disable all

require "json"
require "open3"
require "fileutils"
require "bundler/inline"

puts "Ensuring benchmark peer ecosystem gems are available (bundler/inline)..."
gemfile(true) do
  source "https://rubygems.org"
  gem "caxlsx", "4.5.0"
  gem "xlsxtream", "3.1.0"
  gem "fast_excel", "0.5.0", platform: :mri
  gem "rubyXL", "3.4.38"
  gem "roo", "3.0.0"
  gem "creek", "2.6.3"
  gem "xsv", "1.4.1"
  gem "write_xlsx", "1.15.0"
  gem "simple_xlsx_reader", "5.1.0"
end

RUNS = (ENV["RUNS"] || "3").to_i
ROWS = (ARGV[0] || "100000").to_i
COLS = (ARGV[1] || "10").to_i

AVAILABLE_GEMS = {
  "xlsxrb_stream" => true,
  "xlsxrb_inmemory" => true,
  "xlsxtream" => true,
  "fast_excel" => true,
  "caxlsx" => true,
  "write_xlsx" => true,
  "rubyXL" => true,
  "creek" => true,
  "roo" => true,
  "xsv" => true,
  "simple_xlsx_reader" => true
}

puts "=" * 80
puts "Benchmarking Excel Libraries (#{ROWS} rows x #{COLS} cols = #{ROWS * COLS} cells)"
puts "Runs per benchmark: #{RUNS} (Median reported, Mean calculated)"
puts "Ruby: #{RUBY_DESCRIPTION}"
puts "=" * 80

RUNNER_SCRIPT = <<~'RUBY'
  require "json"
  require "stringio"

  def measure
    gc_before = GC.stat[:count]
    t0 = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    
    yield
    
    t1 = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    gc_after = GC.stat[:count]
    
    # Peak memory in MB via /proc/self/status or getrusage
    peak_mb = 0.0
    if File.exist?("/proc/self/status")
      status = File.read("/proc/self/status")
      if status =~ /VmHWM:\s+(\d+)\s+kB/i
        peak_mb = $1.to_f / 1024.0
      elsif status =~ /VmRSS:\s+(\d+)\s+kB/i
        peak_mb = $1.to_f / 1024.0
      end
    end
    
    {
      time: (t1 - t0),
      peak_memory_mb: peak_mb,
      gc_count: (gc_after - gc_before)
    }
  end

  def generate_row(r, cols)
    base = [r + 1, "User #{r + 1}", 123.45, true, "Active", (r + 1) * 10, "Tokyo", 99.9, false, "Item #{r % 50}"]
    if cols <= base.size
      base.first(cols)
    else
      base + Array.new(cols - base.size) { |c| "col_#{c}_#{r}" }
    end
  end

  lib = ARGV[0]
  mode = ARGV[1] # "write" or "read"
  rows = ARGV[2].to_i
  cols = ARGV[3].to_i
  filename = ARGV[4]

  result = case [lib, mode]
  when ["xlsxtream", "write"]
    require "xlsxtream"
    measure do
      Xlsxtream::Workbook.open(filename) do |wb|
        wb.write_worksheet("Data") do |sheet|
          rows.times do |r|
            sheet << generate_row(r, cols)
          end
        end
      end
    end
  when ["fast_excel", "write"]
    require "fast_excel"
    measure do
      wb = FastExcel.open(filename)
      sheet = wb.add_worksheet("Data")
      rows.times do |r|
        sheet.append_row(generate_row(r, cols))
      end
      wb.close
    end
  when ["caxlsx", "write"]
    require "caxlsx"
    measure do
      p = Axlsx::Package.new
      wb = p.workbook
      wb.add_worksheet(name: "Data") do |sheet|
        rows.times do |r|
          sheet.add_row(generate_row(r, cols))
        end
      end
      p.serialize(filename)
    end
  when ["write_xlsx", "write"]
    require "write_xlsx"
    measure do
      wb = WriteXLSX.new(filename)
      sheet = wb.add_worksheet("Data")
      rows.times do |r|
        sheet.write_row(r, 0, generate_row(r, cols))
      end
      wb.close
    end
  when ["rubyXL", "write"]
    require "rubyXL"
    measure do
      wb = RubyXL::Workbook.new
      sheet = wb[0]
      sheet.sheet_name = "Data"
      rows.times do |r|
        row_data = generate_row(r, cols)
        row_data.each_with_index do |val, c|
          sheet.add_cell(r, c, val)
        end
      end
      wb.write(filename)
    end
  when ["xlsxrb_stream", "write"]
    require_relative "lib/xlsxrb"
    measure do
      Xlsxrb.write(filename) do |wb|
        wb.sheet("Data") do |sheet|
          rows.times do |r|
            sheet.row(generate_row(r, cols))
          end
        end
      end
    end
  when ["xlsxrb_inmemory", "write"]
    require_relative "lib/xlsxrb"
    measure do
      wb = Xlsxrb.build do |b|
        b.sheet("Data") do |s|
          rows.times do |r|
            s.row(generate_row(r, cols))
          end
        end
      end
      Xlsxrb.write(filename, wb)
    end
  when ["xlsxrb_stream", "read"]
    require_relative "lib/xlsxrb"
    measure do
      count = 0
      Xlsxrb.read(filename) do |sheet|
        sheet.each do |row|
          row.cells.each do |cell|
            _val = cell.value
            count += 1
          end
        end
      end
    end
  when ["xlsxrb_inmemory", "read"]
    require_relative "lib/xlsxrb"
    measure do
      wb = Xlsxrb.read(filename)
      count = 0
      wb.sheets.each do |sheet|
        sheet.rows.each do |row|
          row.cells.each do |cell|
            _val = cell.value
            count += 1
          end
        end
      end
    end
  when ["creek", "read"]
    require "creek"
    measure do
      creek = Creek::Book.new(filename)
      count = 0
      creek.sheets.each do |sheet|
        sheet.rows.each do |row|
          row.each_value do |_val|
            count += 1
          end
        end
      end
    end
  when ["roo", "read"]
    require "roo"
    measure do
      xlsx = Roo::Excelx.new(filename)
      count = 0
      xlsx.each_row_streaming do |row|
        row.each do |cell|
          _val = cell&.value
          count += 1
        end
      end
    end
  when ["xsv", "read"]
    require "xsv"
    measure do
      x = Xsv.open(filename)
      count = 0
      x.sheets.each do |sheet|
        sheet.each do |row|
          row.each do |_val|
            count += 1
          end
        end
      end
    end
  when ["simple_xlsx_reader", "read"]
    require "simple_xlsx_reader"
    measure do
      doc = SimpleXlsxReader.open(filename)
      count = 0
      doc.sheets.each do |sheet|
        sheet.rows.each do |row|
          row.each do |_val|
            count += 1
          end
        end
      end
    end
  when ["rubyXL", "read"]
    require "rubyXL"
    measure do
      wb = RubyXL::Parser.parse(filename)
      count = 0
      wb.worksheets.each do |sheet|
        sheet.each do |row|
          next unless row
          row.cells.each do |cell|
            _val = cell&.value
            count += 1
          end
        end
      end
    end
  else
    raise "Unknown benchmark target: #{lib} #{mode}"
  end

  puts result.to_json
RUBY

runner_file = "benchmark_runner.rb"
File.write(runner_file, RUNNER_SCRIPT)

def run_isolated(lib, mode, rows, cols, filename)
  cmd = ["ruby", "-Ilib", "benchmark_runner.rb", lib, mode, rows.to_s, cols.to_s, filename]
  stdout, stderr, status = Bundler.with_unbundled_env do
    Open3.capture3(*cmd)
  end
  unless status.success?
    warn "Failed to run #{lib} #{mode}: #{stderr}"
    return nil
  end
  JSON.parse(stdout.strip, symbolize_names: true)
end

def run_benchmark_series(name, lib, mode, rows, cols, filename, runs)
  unless AVAILABLE_GEMS[lib]
    puts "Skipping #{name} (gem not installed)"
    return nil
  end

  print "Running #{name} (#{runs} runs)... "
  $stdout.flush
  results = []
  runs.times do |_i|
    File.delete(filename) if File.exist?(filename) && mode == "write"
    res = run_isolated(lib, mode, rows, cols, filename)
    if res
      results << res
      print "#{res[:time].round(2)}s "
      $stdout.flush
    else
      print "ERR "
      $stdout.flush
    end
  end
  puts
  return nil if results.empty?

  times = results.map { |r| r[:time] }.sort
  mems = results.map { |r| r[:peak_memory_mb] }.sort
  gcs = results.map { |r| r[:gc_count] }.sort

  median_time = times[times.size / 2]
  mean_time = times.sum / times.size
  median_mem = mems[mems.size / 2]
  median_gc = gcs[gcs.size / 2]

  {
    name: name,
    median_time: median_time,
    mean_time: mean_time,
    median_mem: median_mem,
    median_gc: median_gc
  }
end

# 1. Generate a standard reference file for reading benchmarks
ref_file = "bench_reference_data.xlsx"
puts "\n[Setup] Generating reference file (#{ROWS} x #{COLS}) for read benchmarks..."
run_isolated("xlsxrb_stream", "write", ROWS, COLS, ref_file)

# 2. Benchmark Write
puts "\n=== Benchmarking Write Performance ==="
write_targets = [
  ["xlsxtream 3.1.0", "xlsxtream", "Streaming", "Inline String"],
  ["xlsxrb (Streaming)", "xlsxrb_stream", "Streaming", "SST (Shared)"],
  ["fast_excel 0.5.0 (C)", "fast_excel", "Streaming", "SST (Shared)"],
  ["caxlsx 4.5.0", "caxlsx", "In-Memory", "Inline String"],
  ["write_xlsx 1.15.0", "write_xlsx", "In-Memory", "SST (Shared)"],
  ["xlsxrb (In-Memory)", "xlsxrb_inmemory", "In-Memory", "SST (Shared)"],
  ["rubyXL 3.4.38", "rubyXL", "In-Memory", "Inline String"]
]

write_results = []
write_targets.each do |name, lib, model, storage|
  target_file = "bench_write_#{lib}.xlsx"
  res = run_benchmark_series(name, lib, "write", ROWS, COLS, target_file, RUNS)
  if res
    res[:model] = model
    res[:storage] = storage
    write_results << res
  end
  FileUtils.rm_f(target_file)
end

# 3. Benchmark Read
puts "\n=== Benchmarking Read Performance ==="
read_targets = [
  ["xlsxrb (Streaming)", "xlsxrb_stream", "Streaming"],
  ["xlsxrb (In-Memory)", "xlsxrb_inmemory", "In-Memory"],
  ["creek 2.6.3", "creek", "Streaming"],
  ["roo 3.0.0", "roo", "Streaming"],
  ["simple_xlsx_reader 5.1.0", "simple_xlsx_reader", "Streaming"],
  ["xsv 1.4.1", "xsv", "Streaming"],
  ["rubyXL 3.4.38", "rubyXL", "In-Memory"]
]

read_results = []
read_targets.each do |name, lib, model|
  res = run_benchmark_series(name, lib, "read", ROWS, COLS, ref_file, RUNS)
  if res
    res[:model] = model
    read_results << res
  end
end

# Cleanup
FileUtils.rm_f(ref_file)
FileUtils.rm_f(runner_file)

# Print Tables
puts "\n" + ("=" * 80)
puts "### Write Performance (#{ROWS * COLS} cells: #{ROWS} rows x #{COLS} cols)"
puts ""
puts "| Library                | Model       | Write String Storage | Time (Median) | Time (Mean) | Peak Memory | GC Count |"
puts "| ---------------------- | ----------- | -------------------- | ------------- | ----------- | ----------- | -------- |"
write_results.sort_by { |r| r[:median_time] }.each do |r|
  printf "| %-22s | %-11s | %-20s | %6.2f s      | %6.2f s    | %7.1f MB  | %6.1f   |\n",
         r[:name], r[:model], r[:storage], r[:median_time], r[:mean_time], r[:median_mem], r[:median_gc]
end

puts "\n### Read Performance (#{ROWS * COLS} cells: #{ROWS} rows x #{COLS} cols)"
puts ""
puts "| Library                | Model       | Time (Median) | Time (Mean) | Peak Memory | GC Count |"
puts "| ---------------------- | ----------- | ------------- | ----------- | ----------- | -------- |"
read_results.sort_by { |r| r[:median_time] }.each do |r|
  printf "| %-22s | %-11s | %6.2f s      | %6.2f s    | %7.1f MB  | %6.1f   |\n",
         r[:name], r[:model], r[:median_time], r[:mean_time], r[:median_mem], r[:median_gc]
end
puts "=" * 80
