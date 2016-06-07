require 'fileutils'
module ContentType
  class Content
    def initialize(path = nil)
      path = File.dirname(__FILE__) + "/tempdoc/" if path.nil?
      FileUtils::mkdir_p path unless File.exists?(path)
      @file_path = path + "[Content_Types].xml"
    end

    def create
      ns = {
        "xmlns"=>"http://schemas.openxmlformats.org/package/2006/content-types"
      }
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.Types(ns) {
          xml.Override("PartName"=>"/_rels/.rels" , "ContentType"=>"application/vnd.openxmlformats-package.relationships+xml")
          xml.Override("PartName"=>"/word/_rels/document.xml.rels", "ContentType"=>"application/vnd.openxmlformats-package.relationships+xml")
          xml.Override("PartName"=>"/word/settings.xml" , "ContentType"=>"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")
          xml.Override("PartName"=>"/word/styles.xml" , "ContentType"=>"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
          xml.Override("PartName"=>"/word/document.xml" , "ContentType"=>"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
          xml.Override("PartName"=>"/word/numbering.xml", "ContentType"=>"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
          xml.Override("PartName"=>"/word/fontTable.xml" , "ContentType"=>"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml")
          xml.Override("PartName"=>"/docProps/app.xml" , "ContentType"=>"application/vnd.openxmlformats-officedocument.extended-properties+xml")
          xml.Override("PartName"=>"/docProps/core.xml" , "ContentType"=>"application/vnd.openxmlformats-package.core-properties+xml")
        }
      end

      File.delete(@file_path) if File.exists?(@file_path)
      File.open(@file_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8"?>'))
      end
    end
  end
end
