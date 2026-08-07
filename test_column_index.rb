# frozen_string_literal: true

def column_index(letter)
  str = letter.to_s
  return str.to_i if str.match?(/\A-?\d+\z/)

  str.upcase.chars.reduce(0) { |acc, c| (acc * 26) + (c.ord - "A".ord + 1) } - 1
end

p column_index("0")
p column_index("A")
p column_index("-1")
p column_index("25")
