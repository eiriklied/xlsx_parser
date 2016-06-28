require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe XlsxParser::Xlsx do
  context "instantiating" do
    it "should be possible to instantiate it" do
      doc = XlsxParser::Xlsx.new("data/regular_workbook.xlsx")
      expect(doc).to_not be_nil
    end

    it "should throw a file not found error if the file was not found" do
      expect { XlsxParser::Xlsx.new("non_existent_file") }.to raise_error(Exception)
    end

    it "should throw an error if it was not a valid xlsx document" do
      expect { XlsxParser::Xlsx.new("Gemfile") }.to raise_error(Zip::ZipError)
    end

    it "should read sheet names" do
      doc = XlsxParser::Xlsx.new("data/regular_workbook.xlsx")
      expect(doc.sheets).to eq ['Sheet number one', 'Number two']
    end
  end

end
