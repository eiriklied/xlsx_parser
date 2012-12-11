require 'spec_helper'

describe XlsxParser::Sheet do

  before do
    @shared_strings = ['Cell 1 A', 'Cell 2 B', 'Cell 2 C', 'Dato']
    @sheet = XlsxParser::Sheet.new('test', 'data/sheet1.xml', @shared_strings)
  end

  it 'should answer to cell calls for existing cells' do
    @sheet.cell(1,1).should eq('Cell 1 A')
    @sheet.cell(1,'A').should eq('Cell 1 A')
    
    @sheet.cell(2, 2).should eq('Cell 2 B')
    @sheet.cell(2, 'B').should eq('Cell 2 B')
  end

  it 'should should return nil for non-existing cells' do
    @sheet.cell(-1, -1).should be_nil
    @sheet.cell(0,0).should be_nil
    @sheet.cell(2,1).should be_nil
  end
end