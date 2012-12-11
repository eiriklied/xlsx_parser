require 'zip/zipfilesystem'
require 'nokogiri'

module XlsxParser
  class Xlsx
    def initialize(filename)
      @workbook = Workbook.new(filename)
      @sheets = []
      @last_column = 1 # default value to grow for each sheet
    end

    # An array of sheet names
    def sheets
      @workbook.sheets
    end


  private
    
  end
end