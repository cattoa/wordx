require 'fileutils'
require 'nokogiri'
module Word
  class Styles
    def self.initialize()
      path = File.expand_path(File.dirname(__FILE__) + '/styles/')
      Dir.open(path) do |dir|
        dir.each do |file|
          register_style(path + '/' + file) if file.include?('.yml')
        end
      end
    end

    def self.register_style(style_file_path)
      style = Style.new(YAML.load(IO.read(File.expand_path(style_file_path))))
      if !defined? @@_styles_list
        @@_styles_list = {}
      end
      @@_styles_list[style.name.to_sym] = style
    end

    def self.list
      @@_styles_list.keys
    end

    protected

    def self.method_missing(method, *args)
      if @@_styles_list.keys.include?(method)
        return @@_styles_list[method]
      end
    end
  end

  class Style
    attr_accessor :name, :style, :id, :font_ascii, :font_ansi, :font_east_asia, :font_cs, :font_size, :font_color, :font_lang_default, :font_lang_east_asia,
    :font_lang_bidi, :style_based_on, :style_next, :spacing_line_rule, :spacing_line, :spacing_before, :spacing_after, :indent_left,
    :indent_hanging, :outline_level, :para_keep_with_next, :para_supress_line_numers



    def initialize(style_hash)
      @name = style_hash['name']
      @id = style_hash['id']
      @style = style_hash['style']


      if style_hash.keys.include?('font_ascii')
        @font_ascii = style_hash['font_ascii']
      end


      if style_hash.keys.include?('font_ansi')
        @font_ansi = style_hash['font_ansi']
      end


      if style_hash.keys.include?('font_east_asia')
        @font_east_asia = style_hash['font_east_asia']
      end


      if style_hash.keys.include?('font_cs')
        @font_cs = style_hash['font_cs']
      end


      if style_hash.keys.include?('font_size')
        @font_size = style_hash['font_size']
      end


      if style_hash.keys.include?('font_color')
        @font_color = style_hash['font_color']
      end
      if style_hash.keys.include?('font_lang_default')
        @font_lang_default = style_hash['font_lang_default']
      end
      if style_hash.keys.include?('font_lang_east_asia')
        @font_lang_east_asia = style_hash['font_lang_east_asia']
      end
      if style_hash.keys.include?('font_lang_bidi')
        @font_lang_bidi = style_hash['font_lang_bidi']
      end
      if style_hash.keys.include?('style_based_on')
        @style_based_on = style_hash['style_based_on']
      end
      if style_hash.keys.include?('style_next')
        @style_next = style_hash['style_next']
      end

      if style_hash.keys.include?('spacing_line_rule')
        @spacing_line_rule = style_hash['spacing_line_rule']
      end
      if style_hash.keys.include?('spacing_line')
        @spacing_line = style_hash['spacing_line']
      end
      if style_hash.keys.include?('spacing_before')
        @spacing_before = style_hash['spacing_before']
      end
      if style_hash.keys.include?('spacing_after')
        @spacing_after = style_hash['spacing_after']
      end
      if style_hash.keys.include?('indent_left')
        @indent_left = style_hash['indent_left']
      end
      if style_hash.keys.include?('indent_hanging')
        @indent_hanging = style_hash['indent_hanging']
      end
      if style_hash.keys.include?('outline_level')
        @outline_level = style_hash['outline_level']
      end
      if style_hash.keys.include?('para_keep_with_next')
        @para_keep_with_next = style_hash['para_keep_with_next']
      end
      if style_hash.keys.include?('para_supress_line_numers')
        @para_supress_line_numers = style_hash['para_supress_line_numers']
      end
    end
  end
end
