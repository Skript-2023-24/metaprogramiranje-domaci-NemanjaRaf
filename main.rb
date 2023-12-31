require "google_drive"

session = GoogleDrive::Session.from_config("config.json")
ws = session.spreadsheet_by_key("1q50sps-AQ6wBmAjXctd2Amu97yAWW8javGHIQBz3KQg").worksheets[0]

class Column
  include Enumerable
  attr_reader :values

  def initialize(values)
    @values = values
  end

  def each(&block)
    @values.each do |value|
      block.call(value)
    end
  end

  def sum
    @values.sum
  end

  def avg
    @values.empty? ? 0 : @values.sum.to_f / @values.size
  end

  def method_missing(method_name, *args, &block)
    lookup_value = method_name.to_s
    table.lookup_row_by_column_value(column_name, lookup_value)
  end

  def respond_to_missing?(method_name, include_private = false)
    true
  end
end

class Table
  include Enumerable

  def initialize(data)
    @data = filter_columns(data)
  end

  def filter_columns(data)
    return data if data.empty? || data.first.empty?

    valid_column_indices = data.first.each_index.select do |index|
      !data.first[index].to_s.downcase.match?(/total|subtotal/)
    end

    data.map { |row| valid_column_indices.map { |index| row[index] } }
  end

  def row(index)
    @data[index]
  end

  def each(&block)
    @data.each do |row|
      row.each { |cell| block.call(cell) }
    end
  end

  def [](key)
    if key.is_a?(String)
      col_index = @data.first.index(key)
      return @data.map { |row| row[col_index] } if col_index
    elsif key.is_a?(Numeric)
      return @data[key]
    end
  end

  def []=(key, index, value)
    col_index = @data.first.index(key)
    @data[index][col_index] = value if col_index
  end

  def method_missing(method_name, *args, &block)
    method_string = method_name.to_s
    if method_string.match?(/Kolona/)
      column_name, additional_method = method_string.split(/(?=\.sum|\.avg)/)

      index = @data.first.index { |c| c.downcase == column_name.downcase }
      return super unless index

      column_values = @data[1..-1].map { |row| row[index] }.compact

      numeric_values = column_values.map do |value|
        if value.to_s.match?(/^\d+(\.\d+)?$/)
          value.to_f
        else
          0
        end
      end

      Column.new(numeric_values)
    else
      super
    end
    
  end
  

  def respond_to_missing?(method_name, include_private = false)
    method_name.to_s.end_with?('Kolona', 'sum', 'avg') || super
  end

  def +(other_table)
    if headers_match?(other_table) && same_dimensions?(other_table)
      new_data = [@data.first]

      @data[1..-1].zip(other_table.data[1..-1]).each do |row1, row2|
        new_row = row1.zip(row2).map do |cell1, cell2|
          if numeric?(cell1) && numeric?(cell2)
            cell1.to_f + cell2.to_f
          else
            cell1
          end
        end
        new_data << new_row
      end
      Table.new(new_data)
    else
      raise "Tables do not have matching headers or dimensions!"
    end
  end

  def -(other_table)
    if headers_match?(other_table) && same_dimensions?(other_table)
      new_data = [@data.first]

      @data[1..-1].zip(other_table.data[1..-1]).each do |row1, row2|
        new_row = row1.zip(row2).map do |cell1, cell2|
          if numeric?(cell1) && numeric?(cell2)
            cell1.to_f - cell2.to_f
          else
            cell1
          end
        end
        new_data << new_row
      end
      Table.new(new_data)
    else
      raise "Tables do not have matching headers or dimensions!"
    end
  end

  protected

  attr_reader :data

  private

  def headers_match?(other_table)
    @data.first == other_table.data.first
  end

  def same_dimensions?(other_table)
    @data.length == other_table.data.length && @data.first.length == other_table.data.first.length
  end

  def numeric?(value)
    value.to_s.match?(/^\d+(\.\d+)?$/)
  end
end

data = (1..ws.num_rows).map { |row| (1..ws.num_cols).map { |col| ws[row, col] } }
t = Table.new(data)
t2 = Table.new(data)

puts "Dvodimenzionalni niz: #{t.row(1)}"
puts "Enumerable svih celija: #{t.map { |cell| cell }}"
puts "Pristup koloni preko []: #{t['DrugaKolona']}"
t['PrvaKolona', 1] = 2556
puts "Enumerable svih celija: #{t.map { |cell| cell }}"
puts "Suma i preosek: #{t.drugaKolona.sum}, #{t.prvaKolona.avg}"
puts "Test: #{t.prvaKolona.map { |cell| cell+1 }}"
t3 = t + t2
puts "Razlika dve tabele: #{t3.map { |cell| cell }}"

