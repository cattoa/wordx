require 'fileutils'
require 'nokogiri'
require "wordx/styles"


module Wordx
  class Table
    attr_accessor :row_count, :column_count
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
      @valid_alignments = ["left","right","center","both"]
      @valid_boolean = [true,false]
    end

    def align=(align)
      @align = align if @valid_alignments.include?(align)
    end

    def bold=(bold)
      @bold = bold if @valid_boolean.include?(bold)
    end

  end
end
