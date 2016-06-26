require "wordx/version"
require "wordx/content_type"
require "wordx/docProp"
require "wordx/word"
require "wordx/rels"
require "wordx/paragraphs"



module Wordx
  class Document
    def initialize()
      @@styles = Wordx::Styles.new()
      @@paragraphs = Wordx::Paragraphs.new()
    end

    def list_styles()
      @@styles.list
    end

    def new_paragraph(style=nil,font = nil, font_size=nil)
      style = @@styles.get_style(:Normal) if style.nil?
      font = style.font_ascii if font.nil?
      size = style.font_size() if font_size.nil?
      style.font_ascii = font
      style.font_size = font_size
      para_key = @@paragraphs.new_paragraph(style,font,font_size)
    end

    def add_text(para_key, text = nil)
      text = "" if text.nil?
      @@paragraphs.change_paragraph_text(text,para_key)
    end

    def get_text(para_key)
      return @@paragraphs.get_paragraph_text(para_key)
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
      word.create_document()
    end
  end
end
