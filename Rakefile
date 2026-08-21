# frozen_string_literal: true

require "bundler/gem_tasks"
require "rake/testtask"
require "etc"
require "fileutils"
require "open3"
require "tmpdir"

def dotnet_available?
  system("which dotnet > /dev/null 2>&1")
end

desc "Build the Open XML SDK runner"
task :build_sdk_runner do
  if dotnet_available?
    sh "dotnet build vendor/sdk_runner/sdk_runner.csproj -c Release"
  elsif File.exist?(sdk_runner_dll)
    puts "dotnet not found in PATH, but pre-built sdk_runner.dll exists. Skipping build."
  else
    warn "dotnet command not found and sdk_runner.dll is missing. Cannot build SDK runner."
  end
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
task :ensure_reader_fixtures do
  missing_specs = reader_fixture_specs.reject { |_scenario_name, fixture_path| File.exist?(fixture_path) }
  next if missing_specs.empty?

  Rake::Task[:build_sdk_runner].invoke
  raise "Cannot generate #{missing_specs.size} missing reader fixture(s) because dotnet is not installed." unless dotnet_available?

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

Rake::TestTask.new(test: :ensure_reader_fixtures) do |t|
  t.libs << "test"
  t.libs << "lib"
  t.test_files = FileList["test/xlsxrb/**/*_test.rb", "test/*_test.rb", "test/contract/**/*_test.rb"]
  workers = ENV.fetch("TEST_WORKERS", [Etc.nprocessors, 4].min)
  t.options = "--parallel --n-workers=#{workers}"
end

namespace :test do
  desc "Run tests with runtime type checking enabled (RBS_TEST=1)"
  task :rbs do
    ENV["RBS_TEST"] = "1"
    Rake::Task["test:unit"].invoke
    Rake::Task["test:contract"].invoke
  end
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

  desc "Run property-based tests (PBT)"
  Rake::TestTask.new(:pbt) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/pbt/**/*_test.rb"]
    workers = ENV.fetch("TEST_WORKERS", Etc.nprocessors)
    t.options = "--parallel --n-workers=#{workers}"
  end

  desc "Run memory and performance tests"
  Rake::TestTask.new(:performance) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/performance/**/*_test.rb"]
  end
  desc "Alias for test:performance"
  task perf: :performance

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

require "bundler/audit/task"
Bundler::Audit::Task.new

task default: %i[bundle:audit rubocop typecheck test]

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
  bundle_assets_dir = File.expand_path("docs/wasm/bundle_assets", __dir__)

  # 1. Download original base Wasm if not cached locally
  unless File.exist?(original_wasm_cache)
    puts "Downloading original ruby.wasm for caching..."
    require "open-uri"
    wasm_url = "https://cdn.jsdelivr.net/npm/@ruby/4.0-wasm-wasi@2.9.3-2.9.4/dist/ruby.wasm"
    FileUtils.mkdir_p(File.dirname(original_wasm_cache))
    # rubocop:disable Security/Open
    URI.open(wasm_url) do |stream|
      File.open(original_wasm_cache, "wb") do |file|
        IO.copy_stream(stream, file)
      end
    end
    # rubocop:enable Security/Open
    puts "Original ruby.wasm cached successfully."
  end

  # 2. Recreate and populate the staging bundle directories
  FileUtils.rm_rf(wasm_bundle_dir)
  FileUtils.rm_rf(bundle_assets_dir)
  FileUtils.mkdir_p(wasm_bundle_dir)
  FileUtils.mkdir_p(bundle_assets_dir)

  # A. Write custom static stubs and gateways dynamically
  File.write(File.join(bundle_assets_dir, "openssl.rb"), "# frozen_string_literal: true\n")
  time_rb_path = $LOAD_PATH.lazy.map { |p| File.join(p, "time.rb") }.find { |f| File.exist?(f) }
  FileUtils.cp(time_rb_path, bundle_assets_dir) if time_rb_path

  File.write(File.join(bundle_assets_dir, "opentelemetry.rb"), <<~RUBY)
    # frozen_string_literal: true
    module OpenTelemetry
      def self.tracer_provider
        @tracer_provider ||= Class.new {
          def tracer(*args)
            Class.new {
              def in_span(*args)
                yield Class.new { def record_exception(*args); end; def status=(*args); end }.new
              end
            }.new
          end
        }.new
      end
    end
  RUBY

  # Gateway for REXML
  File.write(File.join(bundle_assets_dir, "rexml.rb"), "# frozen_string_literal: true\nrequire \"rexml/rexml\"\n")

  # Resolve and copy host's rexml files
  rexml_spec_path = $LOAD_PATH.find { |p| File.exist?(File.join(p, "rexml/rexml.rb")) }
  FileUtils.cp_r(File.join(rexml_spec_path, "rexml"), bundle_assets_dir) if rexml_spec_path

  # Gateway for strscan
  File.write(File.join(bundle_assets_dir, "strscan.rb"), <<~RUBY)
    # frozen_string_literal: true
    begin
      require "strscan.so"
    rescue LoadError
    end
    require "strscan/strscan"
  RUBY

  # Resolve and copy host's strscan files
  strscan_spec_path = $LOAD_PATH.find { |p| File.exist?(File.join(p, "strscan/strscan.rb")) }
  if strscan_spec_path
    FileUtils.mkdir_p(File.join(bundle_assets_dir, "strscan"))
    FileUtils.cp(File.join(strscan_spec_path, "strscan/strscan.rb"), File.join(bundle_assets_dir, "strscan/strscan.rb"))
  end

  # B. Copy necessary standard libraries dynamically from host's $LOAD_PATH
  stdlib_files = [
    "date.rb", "delegate.rb", "forwardable.rb", "securerandom.rb",
    "random/formatter.rb", "set.rb", "tempfile.rb", "tmpdir.rb", "fileutils.rb",
    "pp.rb", "prettyprint.rb"
  ]
  stdlib_files.each do |name|
    path = $LOAD_PATH.find { |p| File.exist?(File.join(p, name)) }
    next unless path

    dest_path = File.join(bundle_assets_dir, name)
    FileUtils.mkdir_p(File.dirname(dest_path))
    FileUtils.cp(File.join(path, name), dest_path)
  end

  # C. Patch tmpdir.rb to automatically create /tmp in Wasm virtual filesystem (since Wasm has no writable /tmp by default)
  tmpdir_path = File.join(bundle_assets_dir, "tmpdir.rb")
  if File.exist?(tmpdir_path)
    File.open(tmpdir_path, "a") do |f|
      f.puts "\nclass Dir\n  def self.tmpdir\n    Dir.mkdir(\"/tmp\") rescue nil unless File.directory?(\"/tmp\")\n    \"/tmp\"\n  end\nend\n"
    end
  end

  # D. Copy everything from bundle_assets_dir to wasm_bundle_dir
  FileUtils.cp_r(File.join(bundle_assets_dir, "."), wasm_bundle_dir)

  # E. Copy latest xlsxrb implementation files from lib/
  FileUtils.cp(File.expand_path("lib/xlsxrb.rb", __dir__), File.join(wasm_bundle_dir, "xlsxrb.rb"))
  FileUtils.cp_r(File.expand_path("lib/xlsxrb", __dir__), wasm_bundle_dir)

  # 3. Compile custom ruby.wasm packaging staging libs
  FileUtils.mkdir_p(File.dirname(packed_wasm_path))
  puts "Building packed ruby.wasm from staging bundle..."
  cmd = "bundle exec rbwasm pack #{original_wasm_cache} --dir #{wasm_bundle_dir}::/usr/local/lib/ruby/site_ruby -o #{packed_wasm_path}"
  puts "Executing: #{cmd}"
  raise "Failed to build packed ruby.wasm using rbwasm pack!" unless system(cmd)
end

def download_file(url, dest)
  # Check if the file exists and is not a tiny placeholder/error document
  return if File.exist?(dest) && File.size(dest) > 1024

  puts "Downloading #{url} to #{dest}..."
  FileUtils.mkdir_p(File.dirname(dest))

  require "net/http"
  uri = URI.parse(url)
  temp_dest = "#{dest}.tmp"

  begin
    Net::HTTP.start(uri.host, uri.port, use_ssl: uri.scheme == "https", open_timeout: 15, read_timeout: 90) do |http|
      request = Net::HTTP::Get.new(uri)
      # Specify User-Agent to bypass scraping prevention on CDNs and act as a normal browser
      request["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      # Force raw (identity) encoding to prevent receiving Brotli-compressed (.br) data,
      # which Ruby's Net::HTTP cannot decode automatically, leading to corrupted Wasm files.
      request["Accept-Encoding"] = "identity"

      http.request(request) do |response|
        raise "HTTP error #{response.code}: #{response.message}" if response.code.to_i != 200

        File.open(temp_dest, "wb") do |output|
          response.read_body do |chunk|
            output.write(chunk)
          end
        end
      end
    end
    # Atomic rename to prevent leaving incomplete files on failure
    File.rename(temp_dest, dest)
    puts "Downloaded successfully."

    # Dynamic Brotli Decompression if the server ignored identity encoding and sent Brotli (.br) data
    if File.exist?(dest) && File.binread(dest, 4)&.bytes == [0xCF, 0xFF, 0xFF, 0x7F]
      puts "Detected Brotli compression on #{dest}. Decompressing..."

      unpacked = "#{dest}.unpacked"
      if system("brotli -d -f -o #{unpacked} #{dest}")
        File.rename(unpacked, dest)
        puts "Decompressed #{dest} successfully."
      else
        FileUtils.rm_f(unpacked)
        raise "Failed to decompress Brotli file: #{dest}"
      end
    end
  rescue StandardError => e
    puts "Failed to download #{url}: #{e.message}"
    FileUtils.rm_f(temp_dest)
    FileUtils.rm_f(dest)
    raise "Required asset download failed. Build aborted."
  end
end

def fetch_google_fonts
  css_dest = "docs/fonts/fonts.css"
  return if File.exist?(css_dest)

  puts "Fetching and localizing Google Fonts..."
  font_url = "https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&family=JetBrains+Mono:wght@400;500&display=swap"

  css_content = nil
  begin
    # Specify Chrome User-Agent to ensure Google Fonts returns modern and lightweight .woff2 formats
    # instead of legacy formats (like .ttf or .eot) designed for older browsers
    # rubocop:disable Security/Open
    URI.open(font_url, "User-Agent" => "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36") do |f|
      css_content = f.read
    end
    # rubocop:enable Security/Open
  rescue StandardError => e
    puts "Failed to fetch Google Fonts CSS: #{e.message}"
    return
  end

  urls = css_content.scan(%r{url\((https://fonts\.gstatic\.com/[^)]+)\)}).flatten

  urls.uniq.each do |url|
    filename = url.split("/").last
    local_path = "docs/fonts/#{filename}"
    download_file(url, local_path)
    css_content.gsub!(url, filename)
  end

  FileUtils.mkdir_p("docs/fonts")
  File.write(css_dest, css_content)
  puts "Google Fonts localized successfully."
end

desc "Fetch required external assets for offline usage"
task :fetch_assets do
  require "open-uri"
  require "fileutils"

  # Fetch Ruby WASI JS
  download_file(
    "https://cdn.jsdelivr.net/npm/@ruby/wasm-wasi@2.9.3-2.9.4/dist/browser.umd.js",
    "docs/wasm/browser.umd.js"
  )

  # Fetch ZetaOffice Wasm
  base_zeta_url = "https://cdn.zetaoffice.net/zetaoffice_latest/"
  %w[soffice.js soffice.wasm soffice.data soffice.data.js.metadata qtloader.js].each do |file|
    download_file(base_zeta_url + file, "docs/zetaoffice/#{file}")
  end

  # Fetch and localize Google Fonts
  fetch_google_fonts
end

desc "Generate RDoc documentation including Visual Gallery"
task doc: %i[wasm fetch_assets] do
  FileUtils.rm_rf("doc")

  # RDoc コマンドを実行 (--exclude を指定して docs 配下のプレビュー用アセットの誤パースを回避)
  sh "bundle exec rdoc --op doc " \
     "--exclude 'docs/coi-serviceworker\\.js' " \
     "--exclude 'docs/zeta\\.js' " \
     "--exclude 'docs/office_thread\\.js' " \
     "--exclude 'docs/preview\\.html' " \
     "--title 'xlsxrb Documentation' --main README.md README.md CHANGELOG.md CODE_OF_CONDUCT.md docs/*.md \"docs/visual/VisualGallery.md\" lib/"

  # Copy visual gallery images and files so they are available in RDoc output
  FileUtils.mkdir_p("doc/test/visual/baselines")
  FileUtils.mkdir_p("doc/test/visual/support/illustrations")
  FileUtils.cp_r(Dir.glob("test/visual/baselines/*"), "doc/test/visual/baselines")
  FileUtils.cp_r(Dir.glob("test/visual/support/illustrations/*.png"), "doc/test/visual/support/illustrations")

  FileUtils.mkdir_p("doc/docs/visual/files")
  FileUtils.cp_r(Dir.glob("docs/visual/files/*"), "doc/docs/visual/files")

  FileUtils.mkdir_p("doc/docs/assets")
  FileUtils.cp_r(Dir.glob("docs/assets/*"), "doc/docs/assets")

  # Alias index.html as README_md.html for links referencing README.md
  FileUtils.cp("doc/index.html", "doc/README_md.html") if File.exist?("doc/index.html")

  # --- WebAssembly Playground Integration ---
  puts "Integrating WebAssembly Playground..."

  # Copy playground helper assets and custom Wasm binary to doc directory
  FileUtils.mkdir_p("doc/css")
  FileUtils.mkdir_p("doc/js")
  FileUtils.mkdir_p("doc/wasm")
  FileUtils.mkdir_p("doc/zetaoffice")
  FileUtils.mkdir_p("doc/fonts")

  FileUtils.cp("docs/wasm/wasm_doc_helper.js", "doc/js/wasm_doc_helper.js")
  FileUtils.cp("docs/wasm/wasm_doc_helper.css", "doc/css/wasm_doc_helper.css")
  FileUtils.cp("docs/wasm/ruby.wasm", "doc/wasm/ruby.wasm")
  FileUtils.cp("docs/wasm/browser.umd.js", "doc/wasm/browser.umd.js")
  FileUtils.cp_r(Dir.glob("docs/zetaoffice/*"), "doc/zetaoffice")
  FileUtils.cp_r(Dir.glob("docs/fonts/*"), "doc/fonts")

  # Copy LibreOffice Wasm Preview assets to doc directory
  FileUtils.cp("docs/preview.html", "doc/preview.html")
  FileUtils.cp("docs/coi-serviceworker.js", "doc/coi-serviceworker.js")
  FileUtils.cp("docs/zeta.js", "doc/zeta.js")
  FileUtils.cp("docs/office_thread.js", "doc/office_thread.js")

  # Inject stylesheet and javascript loading tags to all generated HTML docs
  Dir.glob("doc/**/*.html").each do |html_path|
    next if File.basename(html_path) == "preview.html"

    html_content = File.read(html_path)
    depth = html_path.sub(%r{\Adoc/}, "").count("/")
    rel_prefix = "../" * depth

    coi_tag = %(<script src="#{rel_prefix}coi-serviceworker.js"></script>)
    js_tag = %(<script src="#{rel_prefix}js/wasm_doc_helper.js" defer></script>)
    css_tag = %(<link href="#{rel_prefix}css/wasm_doc_helper.css" rel="stylesheet">)

    if html_content.include?("<body")
      modified = html_content.sub("<body", "#{coi_tag}\n#{js_tag}\n#{css_tag}\n<body")
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
      AccessLog: [],
      # If the file is a Brotli-compressed Wasm/Data asset (checked via magic bytes),
      # dynamically inject 'Content-Encoding: br' header so the browser decompresses it natively.
      RequestCallback: lambda { |req, res|
        if req.path.end_with?(".wasm") || req.path.end_with?(".data")
          # Resolve physical file path from req.path manually since res.filename is nil at this stage
          local_path = File.join(File.expand_path("doc", __dir__), req.path)
          if File.exist?(local_path)
            first_bytes = begin
              File.binread(local_path, 4)
            rescue StandardError
              nil
            end
            res["Content-Encoding"] = "br" if first_bytes != "\x00asm"
          end
        end
      }
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
desc "Generate RBS signature files from inline annotations"
task :sig do
  files = FileList["lib/**/*.rb"].to_a
  sh "bundle", "exec", "rbs-inline", "--output=sig/generated", "--base=lib", *files
end

desc "Run static type checking with Steep"
task typecheck: :sig do
  sh "bundle exec steep check"
end
