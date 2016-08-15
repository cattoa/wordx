require "wordx/styles"
require "wordx/table"

module Wordx
  class Paragraphs
    def initialize()
      @current_paragraph = 0
      @paragraphs = {}
    end

    def list
      @paragraphs.keys
    end

    def new_paragraph(style, font, size, bold)
      para_key = get_next_para_key()
      paragraph = Wordx::Paragraph.new(para_key, style, font, size,bold)
      @paragraphs[para_key] =  paragraph
      return paragraph
    end

    def paragraph(key=nil)
      paragraph = get_para(key)
    end

    private

    def get_para(key)
      key = get_current_para_key if key.nil?
      paragraph = @paragraphs[key]
    end

    def get_next_para_key()
      @current_paragraph += 1
      return get_current_para_key()
    end

    def get_current_para_key()
      key = @current_paragraph.to_s.rjust(3,'0')
      key = "para_" + key
      return key.to_sym
    end



  end

  class Paragraph
    attr_accessor :key, :text, :style, :font, :size, :bold


    def initialize(key,style, font, size, bold)
      @key = key
      @text = ""
      @style = style
      @font = font
      @size = size
      @bold = bold
      @tables = {}
      @current_table = 0
    end

    def new_table(row=nil,column=nil)
      key = get_next_table_key
      table = Wordx::Table.new(row,column)
      @tables[key] = table
      return table
    end

    def list_tables
      @tables.keys
    end

    def table(key=nil)
      table = get_table(key)
    end

    private

    def get_table(key)
      key = get_current_table_key if key.nil?
      table = @tables[key]
    end

    def get_next_table_key()
      @current_table += 1
      return get_current_table_key()
    end

    def get_current_table_key()
      key = @current_table.to_s.rjust(3,'0')
      key = "table_" + key
      return key.to_sym
    end


  end
end
