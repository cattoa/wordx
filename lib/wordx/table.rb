require 'fileutils'
require 'nokogiri'
require "wordx/styles"


module Wordx
  class Table
    def initialize(row_count=1,column_count=1)
      @row_count = row_count
      @column_count = column_count
      @current_row = 0
      @current_row = 0
      rows = []
      @table = []
      row = @row_count - 1
      column = @column_count - 1
      for i in (0..row) do
        for j in (0..column) do
          rows.push(Wordx::Cell.new)
        end
        @table.push(rows)
      end
      @valid_alignments = ["left","right","center","both"]
    end
    def length
      return @table.length
    end

    def cell(row,col)
      row = @table[row]
      if !row.nil?
        cell = row[col]
        return cell
      end
    end

    def to_s
      @table.to_s
    end
  end

  class Cell
    attr_accessor :text, :bold, :align

    def initialize()
      @text = ""
      @bold = false
      @align = "left"
    end

    def align(align)
      @align = "left" unless @valid_alignments.include?(align)
    end

    def bold(bold)
      @bold = bold if bold.in?([true,false])
    end

  end
end
