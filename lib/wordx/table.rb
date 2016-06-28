require 'fileutils'
require 'nokogiri'
module Wordx
  class Table
    def initialize(row_count,column_count)
      row_count = 1 unless row_count > 0
      column_count = 1 unless column_count > 0

      @rows = []
      i = 0
      while i < row_count do
        row = Wordx::Row.new()
        @rows.push(row)
        i += 1
      end
      @colums=[]
      i = 0
      while i < column_count do
        row = Wordx::Row.new()
        @columns.push(row)
        i += 1
      end


    end



  end

  class Row

  end

  class Column

  end
end
