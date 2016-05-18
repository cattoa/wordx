require "wordx/version"
require "wordx/content_types"
require "wordx/wordx"

module Wordx
  Wordx::ContentTypes::Content.new().create()

end
