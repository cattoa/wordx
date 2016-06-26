require "wordx/styles"

module Wordx
  class Paragraphs
    def initialize()
      @@current_paragraph = 0
      if !defined? @@paragraphs
        @@paragraphs = {}
      end
    end

    def list
      @@paragraphs.keys
    end

    def new_paragraph(style, font, size)
      para_key = get_next_para_key()
      paragraph = Wordx::Paragraph.new(para_key, style, font, size)
      @@paragraphs[para_key] =  paragraph
      return para_key
    end

    def change_paragraph_text(text, key=nil)
      key = get_current_para_key if key.nil?
      paragraph = @@paragraphs[key]
      unless paragraph.nil?
        text = paragraph.text + text unless paragraph.text.nil?
      end
      paragraph.text = text unless paragraph.nil?
    end

    def change_paragraph_style(style, key=nil)
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

    def get_paragraph_text(key)
      key = get_current_para_key if key.nil?
      paragraph = @@paragraphs[key]
      text = paragraph.text unless paragraph.nil?
      text = "No return" if text.nil?
    end

    private
    def get_next_para_key()
      @@current_paragraph += 1
      return get_current_para_key()
    end

    def get_current_para_key()
      key = @@current_paragraph.to_s.rjust(3,'0')
      key = "para_" + key
      return key.to_sym
    end

  end

  class Paragraph
    attr_accessor :key, :text, :style, :font, :size


    def initialize(key,style, font, size)
      @key = key
      @text = ""
      @style = style
      @font = font
      @size = size
    end

  end
end
