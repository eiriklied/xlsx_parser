class Sheet
  attr_reader :name
  
  def initialize(name, xml_file, shared_strings)
    @name = name
    @doc = parse_sheet(xml_file, shared_strings)
  end

  def cell(row, col)

  end

private

  # parse an xml file and return a 2D 
  # array of the cells in an xlsx sheet
  def parse_sheet(sheet_file, shared_strings)
    xlsx_parser = XlsxSaxParser.new(shared_strings)
    parser = Nokogiri::XML::SAX::Parser.new(xlsx_parser)
    parser.parse(File.read(sheet_file))
    xlsx_parser.doc
  end
end