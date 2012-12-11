require 'spec_helper'

describe Sheet do

  it 'should answer to cell calls' do
    @shared_strings = ['Cell 1 A', 'Cell 2 B', 'Cell 2 C', 'Dato']
    @sheet = Sheet.new('test', 'data/sheet1.xml', @shared_strings)
    @sheet.cell(1,'A')
    @sheet.cell(1,1).should == 'Cell 1 A'
  end
end