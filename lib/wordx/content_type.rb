require 'fileutils'
module ContentType
  class Content
    def initialize(path = File.dirname('#{Rails.root}/tempdoc/')
      FileUtils::mkdir_p path unless File.exists(path)
      @file_path = path + "[Content_Types].xml"
    end

    def create
      File.delete(@file_path) if File.exists?(@file_path)
      File.new(@file_path,"w+") do |fw|
        File.open("#{Rails.root}/lib/wordx/templates/", "r") do |fr|
          fr.each_line do |line|
            fw.write(line)
          end
        end
      end
    end
  end
end
