# frozen_string_literal: true

require "bundler/gem_tasks"
require "rake/testtask"
require "etc"
require "fileutils"
require "open3"
require "tmpdir"

desc "Build the Open XML SDK runner"
task :build_sdk_runner do
  sh "dotnet build vendor/sdk_runner/sdk_runner.csproj -c Release"
end

def reader_fixture_dir
  File.expand_path("test/fixtures/reader_generated", __dir__)
end

def sdk_scenario_dir
  File.expand_path("test/fixtures/sdk_scenarios", __dir__)
end

def sdk_runner_dll
  File.expand_path("vendor/sdk_runner/bin/Release/net8.0/sdk_runner.dll", __dir__)
end

def reader_fixture_specs
  Dir.glob(File.join(sdk_scenario_dir, "reader_*_generated_by_sdk.cs")).map do |scenario_path|
    scenario_name = File.basename(scenario_path, ".cs")
    [scenario_name, File.join(reader_fixture_dir, "#{scenario_name}.xlsx")]
  end
end

def reader_fixture_workers
  Integer(ENV.fetch("READER_FIXTURE_WORKERS", Etc.nprocessors))
rescue ArgumentError
  Etc.nprocessors
end

desc "Ensure SDK-generated reader fixtures exist"
task ensure_reader_fixtures: :build_sdk_runner do
  missing_specs = reader_fixture_specs.reject { |_scenario_name, fixture_path| File.exist?(fixture_path) }
  next if missing_specs.empty?

  FileUtils.mkdir_p(reader_fixture_dir)

  queue = Queue.new
  missing_specs.each { |spec| queue << spec }
  failures = Queue.new

  worker_count = [reader_fixture_workers, missing_specs.size].min
  threads = Array.new(worker_count) do
    Thread.new do
      loop do
        scenario_name, fixture_path = queue.pop(true)
        scenario_path = File.join(sdk_scenario_dir, "#{scenario_name}.cs")
        FileUtils.touch(fixture_path)

        stdout, stderr, status = Open3.capture3(
          "dotnet", sdk_runner_dll, scenario_path, fixture_path
        )

        next if status.success?

        FileUtils.rm_f(fixture_path)
        message = stderr.to_s.strip.empty? ? stdout : stderr
        failures << "Failed to generate reader fixture #{scenario_name}: #{message}"
      rescue ThreadError
        break
      end
    end
  end

  threads.each(&:join)
  next if failures.empty?

  raise failures.pop
end

Rake::TestTask.new(:test) do |t|
  t.libs << "test"
  t.libs << "lib"
  t.test_files = FileList["test/**/*_test.rb"]
  workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
  t.options = "--parallel --n-workers=#{workers}"
end

task test: %i[build_sdk_runner ensure_reader_fixtures]

namespace :test do
  Rake::TestTask.new(:unit) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/xlsxrb/**/*_test.rb", "test/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  Rake::TestTask.new(:contract) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/contract/**/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  Rake::TestTask.new(:e2e) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/e2e/**/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  Rake::TestTask.new(:visual) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/visual/**/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  namespace :fixtures do
    namespace :reader do
      desc "Generate reader fixture XLSX files from SDK scenarios"
      task generate: :ensure_reader_fixtures

      desc "Remove generated reader fixture XLSX files"
      task :clean do
        FileUtils.rm_rf(reader_fixture_dir)
      end
    end
  end
end

task "test:e2e" => %i[build_sdk_runner ensure_reader_fixtures]

require "rubocop/rake_task"

RuboCop::RakeTask.new

task default: %i[test rubocop]

namespace :visual do
  desc "Generate docs/visual/ README.md explanation gallery"
  task :gallery do
    require_relative "test/visual/support/gallery_generator"
    Xlsxrb::Visual::GalleryGenerator.generate
  end

  desc "Generate / update VRT baselines in test/visual/baselines/"
  task :baseline do
    require_relative "test/visual/support/renderer"
    require "securerandom"
    examples = Dir.glob(File.expand_path("examples/visual/*.rb", __dir__))
    baselines_dir = File.expand_path("test/visual/baselines", __dir__)

    FileUtils.rm_rf(baselines_dir)
    FileUtils.mkdir_p(baselines_dir)

    examples.each do |example_path|
      name = File.basename(example_path, ".rb")
      puts "Generating baseline for #{name}..."

      example_tmp_dir = File.join(Dir.tmpdir, "xlsxrb_baselines_build_#{name}_#{SecureRandom.hex(4)}")
      FileUtils.mkdir_p(example_tmp_dir)

      xlsx_path = File.join(example_tmp_dir, "#{name}.xlsx")
      system("ruby", "-Ilib", example_path, xlsx_path)

      dest_dir = File.join(baselines_dir, name)
      begin
        Xlsxrb::Visual::Renderer.render(xlsx_path, dest_dir)
      ensure
        FileUtils.rm_rf(example_tmp_dir)
      end
    end
    puts "Baselines generated successfully."
  end
end

desc "Build custom packed ruby.wasm from staging bundle"
task :wasm do
  original_wasm_cache = File.expand_path("tmp/test_env/ruby.wasm", __dir__)
  packed_wasm_path = File.expand_path("docs/wasm/ruby.wasm", __dir__)
  wasm_bundle_dir = File.expand_path("tmp/wasm_bundle", __dir__)

  # 1. Download original base Wasm if not cached locally
  unless File.exist?(original_wasm_cache)
    puts "Downloading original ruby.wasm for caching..."
    require "open-uri"
    wasm_url = "https://cdn.jsdelivr.net/npm/@ruby/4.0-wasm-wasi@2.9.3-2.9.4/dist/ruby.wasm"
    FileUtils.mkdir_p(File.dirname(original_wasm_cache))
    URI.open(wasm_url) do |stream|
      File.open(original_wasm_cache, "wb") do |file|
        IO.copy_stream(stream, file)
      end
    end
    puts "Original ruby.wasm cached successfully."
  end

  # 2. Synchronize latest lib/xlsxrb.rb to staging bundle directory before packing
  FileUtils.mkdir_p(File.dirname(packed_wasm_path))
  FileUtils.cp(File.expand_path("lib/xlsxrb.rb", __dir__), File.join(wasm_bundle_dir, "xlsxrb.rb"))

  # 3. Compile custom ruby.wasm packaging staging libs
  puts "Building packed ruby.wasm from staging bundle..."
  cmd = "bundle exec rbwasm pack #{original_wasm_cache} --dir #{wasm_bundle_dir}::/usr/local/lib/ruby/site_ruby -o #{packed_wasm_path}"
  puts "Executing: #{cmd}"
  unless system(cmd)
    raise "Failed to build packed ruby.wasm using rbwasm pack!"
  end
end

desc "Generate RDoc documentation including Visual Gallery"
task doc: :wasm do
  FileUtils.rm_rf("doc")
  sh "bundle exec rdoc --title 'xlsxrb Documentation' --main README.md README.md \"docs/visual/VisualGallery.md\" lib/"

  # Copy visual gallery images and files so they are available in RDoc output
  FileUtils.mkdir_p("doc/test/visual/baselines")
  FileUtils.cp_r("test/visual/baselines", "doc/test/visual")
  FileUtils.mkdir_p("doc/test/visual/support/illustrations")
  FileUtils.cp_r(Dir.glob("test/visual/support/illustrations/*.png"), "doc/test/visual/support/illustrations")

  FileUtils.mkdir_p("doc/docs/visual/files")
  FileUtils.cp_r(Dir.glob("docs/visual/files/*"), "doc/docs/visual/files")

  # --- WebAssembly Playground Integration ---
  puts "Integrating WebAssembly Playground..."

  # Copy playground helper assets and custom Wasm binary to doc directory
  FileUtils.mkdir_p("doc/css")
  FileUtils.mkdir_p("doc/js")
  FileUtils.mkdir_p("doc/wasm")
  FileUtils.cp("docs/wasm/wasm_doc_helper.js", "doc/js/wasm_doc_helper.js")
  FileUtils.cp("docs/wasm/wasm_doc_helper.css", "doc/css/wasm_doc_helper.css")
  FileUtils.cp("docs/wasm/ruby.wasm", "doc/wasm/ruby.wasm")

  # Inject stylesheet and javascript loading tags to all generated HTML docs
  Dir.glob("doc/**/*.html").each do |html_path|
    html_content = File.read(html_path)
    depth = html_path.sub(%r{\Adoc/}, "").count("/")
    rel_prefix = "../" * depth

    js_tag = %Q{<script src="#{rel_prefix}js/wasm_doc_helper.js" defer></script>}
    css_tag = %Q{<link href="#{rel_prefix}css/wasm_doc_helper.css" rel="stylesheet">}

    if html_content.include?("<body")
      modified = html_content.sub("<body", "#{js_tag}\n#{css_tag}\n<body")
      File.write(html_path, modified)
    end
  end
  puts "WebAssembly Playground integrated successfully!"
  puts "RDoc documentation generated successfully at doc/"
end

namespace :doc do
  desc "Build and preview RDoc documentation locally"
  task preview: :doc do
    require "webrick"
    
    port = ENV.fetch("PORT", "8000").to_i
    
    # Configure WEBrick to serve 'doc' directory
    server = WEBrick::HTTPServer.new(
      Port: port,
      DocumentRoot: File.expand_path("doc", __dir__),
      Logger: WEBrick::Log.new(nil, WEBrick::BasicLog::WARN),
      AccessLog: []
    )

    puts "=================================================="
    puts " Starting local documentation server at:"
    puts " http://localhost:#{port}/"
    puts " Press Ctrl+C to stop the server"
    puts "=================================================="

    trap("INT") { server.shutdown }

    begin
      server.start
    rescue Errno::EADDRINUSE
      warn "\n[Error] Port #{port} is already in use by another process."
      warn "Please stop the other process or specify a different port."
      warn "Example: PORT=8080 bundle exec rake doc:preview\n\n"
      exit 1
    end
  end
end
