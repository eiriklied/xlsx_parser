require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe Xlsx do
  context "instantiating" do
    it "should be possible to instantiate it" do
      doc = Xlsx.new("data/regular_document.xlsx")
      doc.should_not be_nil
    end
  
    it "should throw a file not found error if the file was not found" do
      expect { Xlsx.new("non_existent_file") }.to raise_error
    end
  end
    
end