# frozen_string_literal: true

require "test_helper"
require "webrick"
require "net/http"

class NetworkStreamingTest < Test::Unit::TestCase
  def setup
    @temp_file = Tempfile.new(["network_streaming", ".xlsx"])
    Xlsxrb.write(@temp_file.path) do |w|
      w.sheet("Data") do |s|
        s.row(%w[Hello World])
      end
    end

    @port = rand(18_080..19_079)
    @server = WEBrick::HTTPServer.new(
      Port: @port,
      Logger: WEBrick::Log.new(File::NULL),
      AccessLog: []
    )

    @server.mount_proc("/") do |_req, res|
      res.chunked = true
      res.body = File.open(@temp_file.path, "rb")
    end

    @server_thread = Thread.new { @server.start }
    sleep 0.1 # Wait for server to start
  end

  def teardown
    @server.shutdown
    @server_thread.join
    @temp_file.close
    @temp_file.unlink
  end

  def test_streaming_read
    uri = URI("http://localhost:#{@port}/")
    Net::HTTP.start(uri.host, uri.port) do |http|
      request = Net::HTTP::Get.new(uri)
      http.request(request) do |response|
        # response is a Net::HTTPResponse, we can use response.enum_for(:read_body)
        enum = response.enum_for(:read_body)

        # Test Xlsxrb.read with the enumerator
        workbook = Xlsxrb.read(enum)

        # Verify it parsed successfully
        assert_equal 1, workbook.sheets.size
      end
    end
  end
end
