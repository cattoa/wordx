module ContentTypes
  class Content
    def initialize(path = "#{Rails.root}/tempdoc/")
      @file_path = path + "[Content_Types].xml"
    end

    def create
      File.open(@file_path,"w+") do |fw|
        File.open("#{Rails.root}/lib/wordx/templates/", "r") do |fr|
          fr.each_line do |line|
            fw.write(line)
          end
        end
      end
    end
  end
end
