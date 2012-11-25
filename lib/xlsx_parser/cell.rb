class Cell
  attr_reader :row, :col, :content

  def initialize(address, content)
    @row, @col = parse_address(address)
    @content = content
  end

  def to_date
    return nil unless @content.is_a?(Integer)
    Date.new(1899, 12, 30) + @content
  end

  private
  def parse_address(address)
    # example address: B2
    /(\D*)(\d*)/ =~ address
    col = $1
    row = $2.to_i

    [row, col]
  end
end