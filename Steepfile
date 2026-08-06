# frozen_string_literal: true

target :lib do
  signature "sig"

  check "lib/xlsxrb.rb"

  library "date"
  library "time"
  library "securerandom"
  library "openssl"
  library "pathname"
  library "tempfile"
  library "bigdecimal"

  repo_path "sig"

  configure_code_diagnostics(Steep::Diagnostic::Ruby.lenient)
end
