require 'spec_helper'

describe SheetSaxParser do

  
  it 'should map shared strings and create Cell objects' do
    @shared_strings = ['Cell 1 A', 'Cell 2 B', 'Cell 2 C', 'Dato']
    xml = File.read('data/sheet1.xml')

    xlsx_parser = SheetSaxParser.new(@shared_strings)
    parser = Nokogiri::XML::SAX::Parser.new(xlsx_parser)
    parser.parse(xml)
    @parsed_doc = xlsx_parser.doc
    @cell_1a = @parsed_doc[0][0]
    @cell_1a.should be_instance_of(Cell)
    @cell_1a.row.should == 1
    @cell_1a.col.should == 'A'
    @cell_1a.content.should == 'Cell 1 A'
    
  end
end