# frozen_string_literal: true

Dir.glob("**/*.{rb,md}").each do |file|
  next if file.start_with?("vendor/")

  content = File.read(file)
  new_content = content.gsub("テスト", "test")
                       .gsub("あ", "a")
  File.write(file, new_content) if content != new_content
end
