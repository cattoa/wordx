require "wordx/version"
require "wordx/content_type"
require "wordx/docProp"
require "wordx/word"
require "wordx/rels"



module Wordx
  class Document
    def initialize()
      @@styles = Wordx::Styles.new()
    end

    def list_styles()
      @@styles.list
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
