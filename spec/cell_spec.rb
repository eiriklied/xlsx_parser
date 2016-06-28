require 'spec_helper'

describe XlsxParser::Cell do
  it 'should parse addresses correctly' do
    @cell = XlsxParser::Cell.new('B2', 'BLABLA')
    expect(@cell.row).to eq 2
    expect(@cell.col).to eq 'B'
  end

  it 'should store content' do
    @cell = XlsxParser::Cell.new('B2', 'BLABLA')
    expect(@cell.content).to eq 'BLABLA'
  end

  it 'should handle dates' do
    @cell = XlsxParser::Cell.new('B2', 41233)
    expect(@cell.to_date).to eq Date.parse('2012-11-20')
  end
end
