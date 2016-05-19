require 'wordx'

describe Wordx::ContentTypes::Content do
  it "create default [Content_Types] file in Rail.root/tempdoc" do
    Wordx::ContentTypes::Content.new().create()
    expect(File.exists("#{Rails.root}/tempdoc/[Content_Types]")).to eq(true)
  end
end
