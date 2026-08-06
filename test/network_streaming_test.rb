require "test_helper"
require "webrick"
require "net/http"
require "thread"

class NetworkStreamingTest < Test::Unit::TestCase
  def setup
    @port = 18080 + rand(1000)
    @server = WEBrick::HTTPServer.new(
      Port: @port,
      Logger: WEBrick::Log.new("/dev/null"),
      AccessLog: []
    )
    
    @server.mount_proc("/") do |req, res|
      file_path = File.expand_path("../test_dates.xlsx", __dir__)
      res.chunked = true
      res.body = File.open(file_path, "rb")
    end
    
    @server_thread = Thread.new { @server.start }
    sleep 0.1 # Wait for server to start
  end

  def teardown
    @server.shutdown
    @server_thread.join
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
