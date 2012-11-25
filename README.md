# XlsxParser

A parser for Excel xlsx documents with a low memory footprint.

xlsx documents are actually just a zip file containing some xml files. The parsers I have used until now are parsing the xml using a DOM parser and if the document gets big, the parser can take up a lot of memory.

This parser reads the document into a simple datastructure in memory using a SAX parser, so the memory usage is quite modest.

## Installation

Add this line to your application's Gemfile:

    gem 'xlsx_parser'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install xlsx_parser

## Usage

This gem aims to provide a subset of the API provided by [roo](https://github.com/hmcgowan/roo)

    # Open a document, parses it and keeps data in memory for fast lookups
    doc = Xlsx.new(@file_import.file.to_file.path) 

    # get the existing sheets for a document
    # as an array of strings
    doc.sheets

    # set the sheet we are working on
    doc.default_sheet = doc.sheets.first

    # get the number for the first row (index starting at 1)
    doc.first_row
    # get the number for the first column (index starting at 1)
    doc.first_column

    # get the number on the last row
    doc.last_row
    # get the number on the last column
    doc.last_column

    # get the value of a cell (returns nil if empty)
    doc.cell(row_num, col_num)


## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
