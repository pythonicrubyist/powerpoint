require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Dashboard
      include Powerpoint::Util

      attr_reader :title, :subtitle, :subtitle_2, :image_path, :image_path_2, :image_name, :data

      def initialize(options={})
        require_arguments [:title, :subtitle, :image_path, :image_path_2, :subtitle_2, :data], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @data = data
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
        render_view('dashboard.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('dashboard.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end