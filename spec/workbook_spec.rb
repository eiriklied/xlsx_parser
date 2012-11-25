require 'spec_helper'

describe Workbook do
  before :each do
    @workbook = Workbook.new('data/regular_workbook.xlsx')
  end

  context 'sheets' do
    it 'should give the names of the sheets in a workbook' do
      @workbook.sheets.should == ['Sheet number one', 'Number two']
    end

    it 'should throw an error if trying to get a sheet that does not exist' do
      expect {
        @workbook.sheet('not a real sheet name')
        }.to raise_error(Workbook::SheetNotFoundException)
    end
  end

  context 'single sheet' do
    it 'should return a sheet by its name' do
      @sheet = @workbook.sheet('Sheet number one')
      @sheet.should be_an_instance_of(Sheet)
      @sheet.name.should == 'Sheet number one'
    end
  end
end