require "wordx/styles"

module Wordx
  class Document
    def initialize()
      @@current_paragraph = -1
      @@paragraphs = []

    end

    def new_paragraph(text, Wordx::Style style, font, size)
      paragraph = Paragragh.new(text, style, font, size)
      @@paragraphs << paragraph
      @@current_paragraph = @@paragraphs.length - 1
    end

    def change_paragraph_text(text)
      paragraph = @@paragraphs[@@current_paragraph]
      paragraph.text(text)
    end

    def change_paragraph_style(Wordx::Style style)
      paragraph = @@paragraphs[@@current_paragraph]
      paragraph.style(style)
    end

    def change_paragraph_font(font)
      paragraph = @@paragraphs[@@current_paragraph]
      paragraph.font(font)
    end

    def change_paragraph_size(size)
      paragraph = @@paragraphs[@@current_paragraph]
      paragraph.size(size)
    end

  end

  class Paragraph
    attr_accessor :text, :style, :font, :size

    def initialize(text, Wordx::Style style, font, size)
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
