require "wordx/version"
require "wordx/content_type"
require "wordx/docProp"
require "wordx/word"
require "wordx/rels"
require "wordx/paragraphs"



module Wordx
  class Document
    def initialize()
      @styles = Wordx::Styles.new()
      @paragraphs = Wordx::Paragraphs.new()
    end

    def list_styles()
      @styles.list
    end

    def get_style(style)
      if style.nil?
        style = @styles.get_style(:DefaultText)
      else
        style = @styles.get_style(style)
      end
    end

    def new_paragraph(style=nil,font = nil, font_size=nil,bold=false)
      if style.nil?
        style = @styles.get_style(:DefaultText)
      else
        style = @styles.get_style(style)
      end
      font = style.font_ascii if font.nil?
      size = style.font_size() if font_size.nil?
      style.font_ascii = font
      style.font_size = font_size
      return @paragraphs.new_paragraph(style,font,font_size,bold)
    end

    def list_paragraphs()
      @paragraphs.list
    end

    def get_paragraph(key = nil)
      @paragraphs.get_paragraph(key)
    end

    def create_document()
      content_type = Wordx::ContentType.new()
      content_type.create()
      doc_props_app = Wordx::DocProp.new()
      doc_props_app.create_app()
      doc_props_app.create_core()
      rels = Wordx::Rels.new()
      rels.create_rels()
      word = Wordx::Word.new()
      word.create_font_table()
      word.create_rels()
      word.create_settings()
      word.create_styles()
      word.create_numbering()
      word.create_document(@paragraphs)
    end
  end
end
