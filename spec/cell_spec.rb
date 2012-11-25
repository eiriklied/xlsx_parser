require 'spec_helper'

describe Cell do
  it 'should parse addresses correctly' do
    @cell = Cell.new('B2', 'BLABLA')
    @cell.row.should == 2
    @cell.col.should == 'B'
  end

  it 'should store content' do
    @cell = Cell.new('B2', 'BLABLA')
    @cell.content.should == 'BLABLA'
  end

  it 'should handle dates' do
    @cell = Cell.new('B2', 41233)
    @cell.to_date.should == Date.parse('2012-11-20')
  end
end