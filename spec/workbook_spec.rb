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

  context 'default_sheet' do
    it 'should default to first sheet if no default sheet has been set' do
      @workbook.default_sheet.name.should eq('Sheet number one')
    end

    it 'should be able to set default sheet sending in a sheet' do
      @workbook.default_sheet = @workbook.sheet(@workbook.sheets.last)
      @workbook.default_sheet.name.should eq('Number two')
    end

    it 'should be able to set default sheet sending in just sheet name' do
      @workbook.default_sheet = 'Number two'
      @workbook.default_sheet.name.should eq('Number two')
    end

    it 'should be able to set default sheet sending in an integer' do
      @workbook.default_sheet = 2
      @workbook.default_sheet.name.should eq('Number two')
      @workbook.default_sheet = 1
      @workbook.default_sheet.name.should eq('Sheet number one')
    end

    it 'should raise an error if trying to default to a sheet name that does not exist' do
      expect { @workbook.default_sheet = 'bogus' }.to raise_error(Workbook::SheetNotFoundException)
    end

    it 'should raise an error if trying to default to a sheet number that does not exist' do
      expect { @workbook.default_sheet = 123 }.to raise_error(Workbook::SheetNotFoundException)
    end

    it 'should raise an error if a bogus param was sent in' do
      expect { @workbook.default_sheet = 0.3 }.to raise_error(TypeError)
    end
  end

  context '*_row and *_column methods' do
    it 'shold resond to *_row methods' do
      @workbook.first_row.should == 1
      @workbook.last_row.should == 4
    end
  
    it 'should respond to *_column methods' do
      @workbook.first_column.should == 1
      @workbook.last_column.should == 3
    end
  end

  
  
end