require 'spec_helper'

describe XlsxParser::Sheet do

  before do
    @shared_strings = ['Cell 1 A', 'Cell 2 B', 'Cell 2 C', 'Dato']
    @sheet = XlsxParser::Sheet.new('test', 'data/sheet1.xml', @shared_strings)
  end

  it 'should answer to cell calls for existing cells' do
    expect(@sheet.cell(1,1)).to eq('Cell 1 A')
    expect(@sheet.cell(1,'A')).to eq('Cell 1 A')

    expect(@sheet.cell(2, 2)).to eq('Cell 2 B')
    expect(@sheet.cell(2, 'B')).to eq('Cell 2 B')
  end

  it 'should return nil for non-existing cells' do
    expect(@sheet.cell(-1, -1)).to be_nil
    expect(@sheet.cell(0,0)).to be_nil
    expect(@sheet.cell(2,1)).to be_nil
  end
end
