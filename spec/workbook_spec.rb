require 'spec_helper'

describe XlsxParser::Workbook do
  before :each do
    @workbook = XlsxParser::Workbook.new('data/regular_workbook.xlsx')
  end

  context 'sheets' do
    it 'should give the names of the sheets in a workbook' do
      @workbook.sheets.should == ['Sheet number one', 'Number two']
    end

    it 'should throw an error if trying to get a sheet that does not exist' do
      expect {
        @workbook.sheet('not a real sheet name')
        }.to raise_error(XlsxParser::Workbook::SheetNotFoundException)
    end
  end

  context 'single sheet' do
    it 'should return a sheet by its name' do
      @sheet = @workbook.sheet('Sheet number one')
      @sheet.should be_an_instance_of(XlsxParser::Sheet)
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
      expect { @workbook.default_sheet = 'bogus' }.to raise_error(XlsxParser::Workbook::SheetNotFoundException)
    end

    it 'should raise an error if trying to default to a sheet number that does not exist' do
      expect { @workbook.default_sheet = 123 }.to raise_error(XlsxParser::Workbook::SheetNotFoundException)
    end

    it 'should raise an error if a bogus param was sent in' do
      expect { @workbook.default_sheet = 0.3 }.to raise_error(TypeError)
    end
  end

  context '*_row and *_column methods' do
    it 'should respond to *_row methods' do
      @workbook.first_row.should == 1
      @workbook.last_row.should == 4
    end

    it 'should respond to *_column methods' do
      @workbook.first_column.should == 1
      @workbook.last_column.should == 3
    end
  end

  context 'cell' do
    it 'should proxy cell calls to sheet' do
      @workbook.cell(1, 'A').should eq('Cell 1 A')

      @workbook.default_sheet = 2
      @workbook.cell(2, 'B').should eq('Sheet 2, cell 2 B')
    end
  end


  it 'should not crash on documents with no sharedStrings.xml' do
      @workbook = XlsxParser::Workbook.new('data/no_shared_strings.xlsx')

      @workbook.cell(1, 'B').should eq('Name')
      @workbook.cell(2, 'B').should eq('Milton Friedman')
  end




end
