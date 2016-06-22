require "wordx/styles"

module Wordx
  class Document
    def initialize()
      @@current_paragraph = 0
      if !defined? @@paragraphs
        @@paragraphs = {}
      end
    end

    def list
      @@paragraphs.keys
    end

    def new_paragraph(text, Wordx::Style style, font, size)
      para_key = get_next_para_key()
      paragraph = Paragragh.new(para_key,text, style, font, size)
      @@paragraphs[para_key] =  paragraph
    end

    def change_paragraph_text(text, key=nil)
      if key.nil?
        key = get_current_para_key
      end
      paragraph = @@paragraphs[key]
      paragraph.text(text) unless paragraph.nil?
    end

    def change_paragraph_style(Wordx::Style style, key=nil)
      if key.nil?
        key = get_current_para_key
      end
      paragraph = @@paragraphs[key]
      paragraph.style(style) unless paragraph.nil?
    end

    def change_paragraph_font(font, key=nil)
      if key.nil?
        key = get_current_para_key
      end
      paragraph = @@paragraphs[key]
      paragraph.font(font) unless paragraph.nil?
    end

    def change_paragraph_size(size, key=nil)
      if key.nil?
        key = get_current_para_key
      end
      paragraph = @@paragraphs[key]
      paragraph.size(size) unless paragraph.nil?
    end

    private
    def get_next_para_key()
      @@current_paragraph++
      next_count = get_current_para()
    end

    def get_current_para_key()
      current_para = "para_" + @@current_paragraph.to_s.rjust(3,'0').to_sym
    end

  end

  class Paragraph
    attr_accessor :key, :text, :style, :font, :size


    def initialize(key, text, Wordx::Style style, font, size)
      @@key = key
      @@text = text
      @@style = style
      @@font = font
      @@size = size
    end

    def text(text_in)
      @@text = text_in
    end

    def style (Wordx::Style style_in)
      @@style = style
    end

    def font(font_in)
      @@font = font_in
    end

    def size(size_in)
      @@size = size_in
    end

  end
end
