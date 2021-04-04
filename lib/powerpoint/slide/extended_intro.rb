require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class ExtendedIntro
      include Powerpoint::Util

      attr_reader :title, :subtitle, :subtitle_2, :coords, :image_path, :image_path_2, :image_name

      def initialize(options={})
        require_arguments [:title, :subtitle, :image_path, :image_path_2, :subtitle_2], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @image_name = File.basename(@image_path)
        @image_name_2 = File.basename(@image_path_2)
      end

      def save(extract_path, index)
        copy_media(extract_path, @image_path)
        copy_media(extract_path, @image_path_2)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def save_rel_xml(extract_path, index)
        render_view('extended_intro_slide_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('extended_intro_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end