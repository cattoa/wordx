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

    def new_paragraph(style, font, size,bold)
      para_key = get_next_para_key()
      paragraph = Wordx::Paragraph.new(para_key, style, font, size,bold)
      @@paragraphs[para_key] =  paragraph
      return para_key
    end

    def set_paragraph_text(text, key=nil)
      paragraph = get_paragraph(key)
      unless paragraph.nil?
        text = paragraph.text + text unless paragraph.text.nil?
      end
      paragraph.text = text unless paragraph.nil?
    end

    def get_paragraph_text(key=nil)
      paragraph = get_paragraph(key)
      text = paragraph.text unless paragraph.nil?
      text = "No return" if text.nil?
      return text
    end

    def set_paragraph_style(style, key=nil)
      paragraph = get_paragraph(key)
      paragraph.style(style) unless paragraph.nil?
    end

    def get_paragraph_style(key=nil)
      paragraph = get_paragraph(key)
      return paragraph.style unless paragraph.nil?
    end

    def set_paragraph_font(font, key=nil)
      paragraph = get_paragraph(key)
      paragraph.font(font) unless paragraph.nil?
    end

    def get_paragraph_font(key=nil)
      paragraph = get_paragraph(key)
      return paragraph.font unless paragraph.nil?
    end

    def set_paragraph_size(size, key=nil)
      paragraph = get_paragraph(key)
      paragraph.size(size) unless paragraph.nil?
    end

    def get_paragraph_size(key=nil)
      paragraph = get_paragraph(key)
      return paragraph.size unless paragraph.nil?
    end

    def set_paragraph_bold(bold=true,key=nil)
      paragraph = get_paragraph(key)
      paragraph.bold = bold
    end

    def get_paragraph_bold(key=nil)
      paragraph = get_paragraph(key)
      return paragraph.bold
    end

    private

    def get_paragraph(key)
      key = get_current_para_key if key.nil?
      paragraph = @@paragraphs[key]
    end

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
    attr_accessor :key, :text, :style, :font, :size, :bold


    def initialize(key,style, font, size, bold)
      @key = key
      @text = ""
      @style = style
      @font = font
      @size = size
      @bold = bold
    end

  end
end
