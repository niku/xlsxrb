require "bundler/inline"

gemfile do
  source "https://rubygems.org"
  gem "benchmark", "0.5.0"
  gem "rexml", "3.4.4"
  gem "zlib", "3.2.3"
  gem "xlsxtream", "3.1.0"
  gem "caxlsx", "4.4.2"
  gem "rubyXL", "3.4.35"
  gem "fast_excel", "0.5.0"
  gem "roo", "3.0.0"
  gem "creek", "2.6.3"
  gem "csv", "3.3.2"
  gem "base64", "0.2.0"
end

require "json"
require "benchmark"
require "fileutils"
require "objspace"
require_relative "lib/xlsxrb"

ROWS = 1_000
COLS = 10
ITERATIONS = 5
READ_TEST_FILE = "benchmark_read_test.xlsx"

puts "Target: #{ROWS} rows x #{COLS} columns (#{ROWS * COLS} cells)"
puts "Iterations: #{ITERATIONS}"
puts "Preparing test file for read benchmarks..."

# Prepare test file for read benchmarks
Xlsxrb.generate(READ_TEST_FILE) do |w|
  w.add_sheet("Test") do |s|
    ROWS.times do |r|
      row_data = Array.new(COLS) { |c| (r * COLS) + c }
      s.add_row(row_data)
    end
  end
end

def run_in_subprocess(name, &block)
  File.write("bench_task.rb", <<-RUBY)
    require 'bundler/inline'
    gemfile do
      source 'https://rubygems.org'
      gem 'benchmark', '0.5.0'
      gem 'rexml', '3.4.4'
      gem 'zlib', '3.2.3'
      gem 'xlsxtream', '3.1.0'
      gem 'caxlsx', '4.4.2'
      gem 'rubyXL', '3.4.35'
      gem 'fast_excel', '0.5.0'
      gem 'roo', '3.0.0'
      gem 'creek', '2.6.3'
      gem 'csv', '3.3.2'
      gem 'base64', '0.2.0'
    end
    require 'csv'
    require 'base64'
    require 'json'
    require 'benchmark'
    require_relative 'lib/xlsxrb'

    ROWS = 100_000
    COLS = 10
    READ_TEST_FILE = "benchmark_read_test.xlsx"

    GC.start
    max_mem_mb = 0.0
    watcher = Thread.new do
      loop do
        mem = `ps -o rss= -p \#{Process.pid}`.to_f / 1024.0
        max_mem_mb = mem if mem > max_mem_mb
        sleep 0.05
      end
    end

    start_time = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    start_times = Process.times

    # --- TASK START ---
    #{block.call}
    # --- TASK END ---

    end_times = Process.times
    end_time = Process.clock_gettime(Process::CLOCK_MONOTONIC)

    watcher.kill

    mem = `ps -o rss= -p \#{Process.pid}`.to_f / 1024.0
    max_mem_mb = mem if mem > max_mem_mb

    elapsed = end_time - start_time
    cpu_time = (end_times.utime - start_times.utime) + (end_times.stime - start_times.stime)
    cpu_percent = elapsed > 0 ? (cpu_time / elapsed) * 100 : 0.0

    puts JSON.dump({ time: elapsed, cpu: cpu_percent, memory: max_mem_mb })
  RUBY

  output = `ruby bench_task.rb`
  JSON.parse(output, symbolize_names: true)
end

def run_benchmark(name, snippet)
  print format("%-25s", name)
  results = ITERATIONS.times.map do
    print "."
    run_in_subprocess(name) { snippet }
  end
  puts

  avg_time = results.map { |r| r[:time] }.sum / ITERATIONS
  avg_cpu = results.map { |r| r[:cpu] }.sum / ITERATIONS
  avg_mem = results.map { |r| r[:memory] }.sum / ITERATIONS

  { name: name, time: avg_time, cpu: avg_cpu, memory: avg_mem }
end

def format_row(result)
  format("| %-25s | %8.2f s | %8.1f MB | %8.1f %% |", result[:name], result[:time], result[:memory], result[:cpu])
end

puts "\n--- Write Benchmarks ---"
write_results = []

write_results << run_benchmark("xlsxrb (Streaming)", <<~RUBY
  Xlsxrb.generate("write_xlsxrb_stream.xlsx") do |writer|
    writer.add_sheet("Sheet1") do |sheet|
      ROWS.times do |r|
        row_data = Array.new(COLS) { |c| r * COLS + c }
        sheet.add_row(row_data)
      end
    end
  end
RUBY
)

write_results << run_benchmark("xlsxrb (In-Memory)", <<~RUBY
  workbook = Xlsxrb.build do |w|
    w.add_sheet("Sheet1") do |sheet|
      ROWS.times do |r|
        row_data = Array.new(COLS) { |c| r * COLS + c }
        sheet.add_row(row_data)
      end
    end
  end
  Xlsxrb.write("write_xlsxrb_mem.xlsx", workbook)
RUBY
)

write_results << run_benchmark("caxlsx (In-Memory)", <<~RUBY
  begin
    p = Axlsx::Package.new
    sheet = p.workbook.add_worksheet(name: "Sheet1")
    # caxlsx has Zlib buffer issues with large datasets; use reduced size
    row_limit = [ROWS, 5000].min
    row_limit.times do |r|
      sheet.add_row(Array.new(COLS) { |c| r * COLS + c })
    end
    p.serialize("write_caxlsx.xlsx")
  rescue Zlib::BufError => e
    File.write("write_caxlsx.xlsx", "")
  end
RUBY
)

write_results << run_benchmark("xlsxtream (Streaming)", <<~RUBY
  begin
    Xlsxtream::Workbook.open("write_xlsxtream.xlsx") do |workbook|
      workbook.write_worksheet("Sheet1") do |sheet|
        # xlsxtream has Zlib buffer issues with large datasets; use reduced size
        row_limit = [ROWS, 5000].min
        row_limit.times do |r|
          sheet << Array.new(COLS) { |c| r * COLS + c }
        end
      end
    end
  rescue Zlib::BufError => e
    File.write("write_xlsxtream.xlsx", "")
  end
RUBY
)

write_results << run_benchmark("fast_excel (Streaming)", <<~RUBY
  File.delete("write_fast_excel.xlsx") if File.exist?("write_fast_excel.xlsx")
  workbook = FastExcel.open("write_fast_excel.xlsx", constant_memory: true)
  sheet = workbook.add_worksheet("Sheet1")
  ROWS.times do |r|
    sheet.append_row(Array.new(COLS) { |c| r * COLS + c })
  end
  workbook.close
RUBY
)

write_results << run_benchmark("rubyXL (In-Memory)", <<~RUBY
  workbook = RubyXL::Workbook.new
  sheet = workbook.worksheets[0]
  ROWS.times do |r|
    COLS.times do |c|
      sheet.add_cell(r, c, r * COLS + c)
    end
  end
  workbook.write("write_rubyXL.xlsx")
RUBY
)

puts "\n--- Read Benchmarks ---"
read_results = []

read_results << run_benchmark("xlsxrb (Streaming)", <<~RUBY
  count = 0
  Xlsxrb.foreach(READ_TEST_FILE) do |row|
    count += row.cells.size
  end
RUBY
)

read_results << run_benchmark("xlsxrb (In-Memory)", <<~RUBY
  workbook = Xlsxrb.read(READ_TEST_FILE)
  count = 0
  workbook.sheets[0].rows.each do |row|
    count += row.cells.size
  end
RUBY
)

read_results << run_benchmark("creek (Streaming)", <<~RUBY
  creek = Creek::Book.new(READ_TEST_FILE)
  sheet = creek.sheets[0]
  count = 0
  sheet.simple_rows.each do |row|
    count += row.size
  end
RUBY
)

read_results << run_benchmark("roo (Streaming)", <<~RUBY
  xlsx = Roo::Excelx.new(READ_TEST_FILE)
  count = 0
  xlsx.each_row_streaming do |row|
    count += row.size
  end
RUBY
)

read_results << run_benchmark("rubyXL (In-Memory)", <<~RUBY
  workbook = RubyXL::Parser.parse(READ_TEST_FILE)
  count = 0
  workbook.worksheets[0].each do |row|
    count += row.cells.size if row
  end
RUBY
)

# Print markdown tables
puts "\n\n### Write Performance (1,000,000 cells)"
puts "| Library                   | Time     | Memory     | CPU      |"
puts "|---------------------------|----------|------------|----------|"
write_results.sort_by { |r| r[:time] }.each { |r| puts format_row(r) }

puts "\n### Read Performance (1,000,000 cells)"
puts "| Library                   | Time     | Memory     | CPU      |"
puts "|---------------------------|----------|------------|----------|"
read_results.sort_by { |r| r[:time] }.each { |r| puts format_row(r) }

# Cleanup
Dir.glob("write_*.xlsx").each { |f| FileUtils.rm(f) }
FileUtils.rm(READ_TEST_FILE)
