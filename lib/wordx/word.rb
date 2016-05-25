require 'fileutils'
require 'nokogiri'
module Word
  class Content
    def initialize(path = nil)
      path = File.dirname(__FILE__) + "/tempdoc/word/" if path.nil?
      path_rels = path + "/_rels/"
      FileUtils::mkdir_p path unless File.exists?(path)
      @word_path = path
      FileUtils::mkdir_p path_rels unless File.exists?(path_rels)
      @word_rels_path = path_rels
    end

    def create_font_table(fonts = nil)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:xsi"=>"http://www.w3.org/2001/XMLSchema-instance"
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
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end
    end

    def create_rels(fonts = nil)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:xsi"=>"http://www.w3.org/2001/XMLSchema-instance"
      }
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.Relationships(ns) {
          xml.Relationship "Id"=>"rId1", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "Target"=>"styles.xml"
          xml.Relationship "Id"=>"rId2", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "Target"=>"numbering.xml"
          xml.Relationship "Id"=>"rId3", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable", "Target"=>"fontTable.xml"
          xml.Relationship "Id"=>"rId4", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "Target"=>"settings.xml"
        }
      end

      word_rels_path = @word_rels_path + "document.xml.rels"
      File.delete(word_rels_path) if File.exists?(word_rels_path)
      File.open(word_rels_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end

    end

    def create_settings(fonts = nil)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      }

      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:w].settings(ns) {
          xml[:w].zoom "w:percent"=> "100"
          xml[:w].defaultTabStop "w:val"=>"700"
          xml[:w].compat
          xml[:w].themeFontLang "w:val"=> "", "w:eastAsia" => "", "w:bidi"=> ""
        }
      end

      word_path = @word_path + "settings.xml"
      File.delete(word_path) if File.exists?(word_path)
      File.open(word_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end

    end

    def create_document(doc = nil)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:mc"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      }
      if doc.nil?
        doc = [{:style=>"Heading1",:text=>"Heading 1 Text",:before=>"240",:after=>"120"}]
        doc << {:style=>"Heading2",:text=>"Heading 2 Text"}
        doc << {:style=>"Heading3",:text=>"Heading 3 Text"}
        doc << {:style=>"Heading",:text=>"Heading Text"}
        doc << {:style=>"Normal",:text=>"Normal Texy"}
      end


      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:w].document(ns,"mc:Ignorable"=>"w14 wp14") {
          doc.each do |doc_param|
            xml[:w].body {
              xml[:w].p{
                xml[:w].pPr {
                  xml[:w].pStyle("w:val"=>doc_param[:style]) unless doc_param[:style].nil?
                  spacing_detail = {"w:before"=>doc_param[:before],
                                    "w:after"=>doc_param[:after]}
                  xml[:w].spacing (spacing_detail)
                  xml[:w].rPr
                }
                xml[:w].r{
                  xml[:w].rPr
                  xml[:w].t (doc_param[:text]) unless doc_param[:text].nil?
                }
              }
            }
          end
        }
      end
      doc_path = @word_path + "document.xml"
      File.delete(doc_path) if File.exists?(doc_path)
      File.open(doc_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end
    end

    def create_styles(doc_styles = nil)

      if doc_styles.nil?
        doc_styles = [{:type=>"paragraph",:styleId=>"Normal", :name=>"Normal", :basedOn=>"", :next=>"Normal", :jc=>"left", :ascii=>"Liberation Serif", :size=>"24",:color=> "00000A"}]
      end
      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w14"=>"http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:mc"=>"http://schemas.openxmlformats.org/markup-compatibility/2006"
      }

      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:w].styles(ns,"mc:Ignorable"=>"w14") {
          xml[:w].docDefaults{
            xml[:w].rPrDefault {
              xml[:w].rPr {
                xml[:w].rFonts
                xml[:w].szCs
                xml[:w].lang
              }
            }
            xml[:w].pPrDefault {
              xml[:w].pPr {
                xml[:w].widowControl
              }
            }
          }
          doc_styles.each do |doc_style|
            style_attr = {"w:type"=>doc_style[:type],"w:styleId"=>doc_style[:styleId]}
            xml[:w].style(style_attr)  {
              xml[:w].name "w:va"=>doc_style[:name] unless doc_style[:name].nil?
              xml[:w].basedOn
              xml[:w].next
              xml[:w].qFormat
              xml[:w].pPr {
                xml[:w].keepNext
                xml[:w].spacing
                xml[:w].widowControl
                xml[:w].bidi
                xml[:w].jc
              }
              xml[:w].rPr {
                xml[:w].rFonts
                xml[:w].color
                xml[:w].sz
                xml[:w].szCs
                xml[:w].lanf
              }
            }
          end
        }
      end

      doc_path = @word_path + "styles.xml"
      File.delete(doc_path) if File.exists?(doc_path)
      File.open(doc_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end
    end

  end
end
