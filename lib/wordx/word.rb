require 'fileutils'
require 'nokogiri'
module Word
  class Content
    def initialize(path = nil)
      path = File.dirname(__FILE__) + "/tempdoc/word/" if path.nil?
      FileUtils::mkdir_p path unless File.exists?(path)
      @word_path = path
    end
  end
end
