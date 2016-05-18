require 'wordx'

describe Wordx::ContentTypes::Content do
  it "create default [Content_Types] file in Rail.root/tempdoc"
    Wordx::ContentTypes::Content.new().create()
    (File.exists("#{Rails.root}/tempdoc/[Content_Types]").is_true
end
