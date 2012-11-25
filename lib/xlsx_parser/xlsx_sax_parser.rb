#require 'rubygems'
require 'nokogiri'


class XlsxSaxParser < Nokogiri::XML::SAX::Document
  attr_reader :doc
  
  def initialize(shared_strings)
    raise "Need some shared strings to populate the sheet!" unless shared_strings
    @shared_strings = shared_strings
    @buffer = ''
    @doc = []
  end

  def start_element name, attrs = []
    start_row    name, attrs
    start_column name, attrs
  end

  def end_element name
    end_row name
    end_column name
  end

  # callbacks that can be called multiple times
  # filling a buffer with element contents
  def characters(string)
    @buffer << string
  end
  def cdata_block(string)
    characters(string)
  end

  def cell row, col
    @doc[row-1][col-1]
  end

  private

  def start_row name, attrs
    return unless name == 'row'
    
    # in the xml empty rows might not be present.
    # here we find empty rows we might have missed and 
    # populate the doc with an empty array (which is an empty row)
    row_num = attribute('r', attrs).to_i
    while @doc.length+1 < row_num
      @doc << []
    end
    
    @current_row = []
  end

  def end_row name
    return unless name == 'row'
    @doc << @current_row
  end

  def end_column name
    return unless name == 'c'
  
    if @is_string
      @current_row << @shared_strings[@buffer.to_i]
    else
      @current_row << @buffer
    end
  end

  def start_column name, attrs
    return unless name == 'c'
    @buffer = ''
    # flag to see if we need to get the value from a string pool
    @is_string = (attribute('t', attrs) == 's')
  end

  def parse_col_and_row(str)
    arr = str.scan(/([A-Za-z]{1,2})(\d*)/)
    arr.first
  end

  def attribute(find, arr)
    arr.each do |elem|
      return elem[1] if elem[0] == find
    end
    nil
  end
end
