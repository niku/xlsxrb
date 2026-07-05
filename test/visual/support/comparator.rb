# frozen_string_literal: true

require "open3"

module Xlsxrb
  module Visual
    class Comparator
      # Compares two images using ImageMagick's `compare` command.
      # Returns the number of differing pixels.
      # Writes a diff image highlight differences to diff_path.
      def self.compare(baseline_path, candidate_path, diff_path, fuzz: "5%")
        # ImageMagick compare outputs AE metric value to stderr.
        # Status code is 0 (match), 1 (mismatch), or 2 (error).
        stdout, stderr, status = Open3.capture3(
          "compare",
          "-metric", "AE",
          "-fuzz", fuzz,
          baseline_path,
          candidate_path,
          diff_path
        )

        raise "ImageMagick compare command failed: #{stderr}\n#{stdout}" if status.exitstatus == 2

        # Parse number of differing pixels from stderr
        stderr.to_s.strip.to_i
      end
    end
  end
end
