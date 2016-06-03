require "wordx/version"
require "wordx/content_type"
require "wordx/docProp"
require "wordx/word"



module Wordx
  styles = Wordx::Styles.new()
  content_type = ContentType::Content.new()
  content_type.create()
  doc_props_app = DocProp::Content.new()
  doc_props_app.create_app()
  doc_props_app.create_core()
  word = Word::Content.new()
  word.create_font_table()
  word.create_rels()
  word.create_settings()
  word.create_styles()
end
