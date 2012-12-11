class Sheet
  attr_reader :name
  
  def initialize(name, xml_file, shared_strings)
    @name = name
    @sheet = read_sheet(xml_file, shared_strings)
    #@sheet = read_cells_into_2d_structure(sheet_cells)
  end

  def cell(row, col)
    row_num, col = normalize(row, col)
    return nil unless @sheet[row_num-1]

    selected_cell = @sheet[row_num-1].select{ |cell| cell.col == col }.first
    return nil unless selected_cell
    selected_cell.content
  end

  def first_row
    return 1 unless @sheet.first
    @sheet.first.first.row
  end

  def last_row
    return 1 unless @sheet.last
    @sheet.last.first.row
  end



  def first_column
    return @first_column if @first_column
    @first_column = @sheet.flatten.map{|cell| Sheet.letter_to_number(cell.col) }.min
  end

  def last_column
    return @last_column if @last_column
    @last_column = @sheet.flatten.map{|cell| Sheet.letter_to_number(cell.col) }.max
  end

private


  # parse an xml file and return array of cells
  def read_sheet(sheet_file, shared_strings)
    xlsx_parser = SheetSaxParser.new(shared_strings)
    parser = Nokogiri::XML::SAX::Parser.new(xlsx_parser)
    parser.parse(File.read(sheet_file))
    xlsx_parser.doc
  end


  #
  # grabbed from https://github.com/hmcgowan/roo/blob/master/lib/roo/generic_spreadsheet.rb
  #
  # converts cell coordinate to numeric values of row,col
  def normalize(row,col)
    if row.class == String
      if col.class == Fixnum
        # ('A',1):
        # ('B', 5) -> (5, 2)
        row, col = col, row
      else
        raise ArgumentError
      end
    end
    if col.class == Fixnum
      col = Sheet.number_to_letter(col)
    end
    return row,col
  end

  # convert letters like 'AB' to a number ('A' => 1, 'B' => 2, ...)
  def self.letter_to_number(letters)
    result = 0
    while letters && letters.length > 0
      character = letters[0,1].upcase
      num = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".index(character)
      raise ArgumentError, "invalid column character '#{letters[0,1]}'" if num == nil
      num += 1
      result = result * 26 + num
      letters = letters[1..-1]
    end
    result
  end

  # convert a number to something like 'AB' (1 => 'A', 2 => 'B', ...)
  def self.number_to_letter(n)
    letters=""
    while n > 0
      num = n%26
      letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[num-1,1] + letters
      n = n.div(26)
    end
    letters
  end

end