require 'fileutils'
module ContentType
  class Content
    def initialize(path = nil)
      path = File.dirname(__FILE__) + "/tempdoc/" if path.nil?
      FileUtils::mkdir_p path unless File.exists?(path)
      @file_path = path + "[Content_Types].xml"
    end

    def create
      File.delete(@file_path) if File.exists?(@file_path)
      File.open(@file_path, "w+") do |fw|
        File.open(File.dirname(__FILE__) + "/templates/[Content_Types].xml", "r") do |f|
          f.each_line do |line|
            fw.write(line)
          end
        end
      end
    end
  end
end
