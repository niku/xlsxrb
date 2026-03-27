# Xlsxrb

A Ruby library for reading and writing XLSX files with streaming support.

## Motivation

The Ruby ecosystem already has great XLSX libraries. Each is well-designed for its purpose:

| Library | Read | Write | Streaming (low memory) |
|---------|------|-------|------------------------|
| [roo](https://rubygems.org/gems/roo) | ✅ | ❌ | ❌ |
| [creek](https://rubygems.org/gems/creek) | ✅ | ❌ | ✅ |
| [caxlsx / axlsx](https://rubygems.org/gems/caxlsx) | ❌ | ✅ | ❌ |
| [xlsxtream](https://rubygems.org/gems/xlsxtream) | ❌ | ✅ | ✅ |
| [rubyXL](https://rubygems.org/gems/rubyXL) | ✅ | ✅ | ❌ |
| [fast_excel](https://rubygems.org/gems/fast_excel) | ❌ | ✅ | ✅ |

If you need to read large files efficiently, [creek](https://rubygems.org/gems/creek) is a great choice. If you need to write large files, [xlsxtream](https://rubygems.org/gems/xlsxtream) does that well. These libraries make deliberate tradeoffs, and they do so thoughtfully.

`xlsxrb` is for cases where you need **both** reading and writing in a single library, while also keeping memory usage predictable for large files.

## Installation

```bash
bundle add xlsxrb
```

Or without Bundler:

```bash
gem install xlsxrb
```

## Usage

TODO: Write usage instructions here

## Specification

This project aims to be compliant with [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) (Office Open XML file formats). Specifically, the library targets the **Transitional** version of the specification rather than the **Strict** version. The Transitional version (detailed in Part 4) is the format most commonly produced and consumed by existing spreadsheet applications, making it the practical choice for real-world interoperability.

For reference, the following specification files from the Ecma International website are located in the `vendor/docs/` directory:

- `vendor/docs/ECMA-376-Part1/Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf`: Part 1 - Fundamentals And Markup Language Reference
- `vendor/docs/ECMA-376-Part2/ECMA-376-2_5th_edition_december_2021.pdf`: Part 2 - Open Packaging Conventions
- `vendor/docs/ECMA-376-Part3/ECMA-376-3_5th_edition_december_2015.pdf`: Part 3 - Markup Compatibility and Extensibility
- `vendor/docs/ECMA-376-Part4/Ecma Office Open XML Part 4 - Transitional Migration Features.pdf`: Part 4 - Transitional Migration Features

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/niku/xlsxrb. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/niku/xlsxrb/blob/main/CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the Xlsxrb project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/niku/xlsxrb/blob/main/CODE_OF_CONDUCT.md).
