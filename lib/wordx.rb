require "wordx/version"
require "wordx/content_type"
require "wordx/docProp"



module Wordx
  content_type = ContentType::Content.new()
  content_type.create()
  doc_props_app = DocProp::Content.new()
  doc_props_app.create_app()
  doc_props_app.create_core()

end
