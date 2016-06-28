require 'fileutils'
require 'nokogiri'
require "wordx/styles"


module Wordx
  class Table
    attr_accessor :cell
    def initialize(row_count=1,column_count=1,style=nil)
      @current_table = 0
      row_count = 1 unless row_count > 0
      column_count = 1 unless column_count > 0
      @style = style unless style.nil?

      @cells=[]
      i = 0
      j=0
      while i < row_count do
        while j < column_count do
          cell = Wordx::Cell.new(i,j)
          @cells[i,j] = cell
          j += 1
        end
        i += 1
      end

    end
  end

  class Cell
    attr_accessor :row, :column, :text, :style

    def initialize(row = 0, coulmn = 0, style_key = nil)
      styles = Wordx::Styles.new()
      if style_key.nil?
        style = styles.get_style(:DefaultText)
      else
        style = styles.get_style(style_key)
      end
      @row = row
      @column = column
    end

  end

  class Row

  end

  class Column

  end
end
