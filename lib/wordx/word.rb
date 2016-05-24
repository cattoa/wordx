require 'fileutils'
require 'nokogiri'
module Word
  class Content
    def initialize(path = nil)
      path = File.dirname(__FILE__) + "/tempdoc/word/" if path.nil?
      FileUtils::mkdir_p path unless File.exists?(path)
      @word_path = path
    end

    def create_font_table(fonts = nil)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      }
      if fonts.nil?
        fonts = [{:name=>"Times New Roman",:charset=>"00",:family=>"roman",:pitch=>"variable"}]
        fonts << {:name=>"Symbol",:charset=>"02",:family=>"roman",:pitch=>"variable"}
        fonts << {:name=>"Arial",:charset=>"00",:family=>"swiss",:pitch=>"variable"}
        fonts << {:name=>"Liberation Serif",:altName =>"Times New Roman",:charset=>"01",:family=>"roman",:pitch=>"variable"}
        fonts << {:name=>"Liberation Sans",:altName =>"Times New Roman",:charset=>"01",:family=>"roman",:pitch=>"variable"}
      end


      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:w].fonts(ns) {
          fonts.each do |font_array|
            puts font_array
              fn = {
                "w:name" =>font_array[:name]
              }
            xml[:w].font(fn) {
              xml[:w].charset("w:val"=>font_array[:charset])
              xml[:w].altName("w:val"=>font_array[:altName]) unless font_array[:altName].nil?
              xml[:w].family("w:val"=>font_array[:family])
              xml[:w].pitch("w:val"=>font_array[:pitch])
            }
          end
        }
      end
      word_path = @word_path + "fontTable.xml"
      File.delete(word_path) if File.exists?(word_path)
      File.open(word_path, "w+") do |fw|
        fw.write(builder.to_xml)
      end
    end
  end
end
