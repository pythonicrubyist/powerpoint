require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Dashboard
      include Powerpoint::Util

      attr_reader :title, :subtitle, :graph_1_title, :graph_1_subtitle, :graph_2_title, :graph_2_subtitle, :image_path, :image_path_2,:image_path_3,:image_path_4, :image_name, :logo, :data

      def initialize(options={})
        require_arguments [:title, :subtitle, :graph_1_title, :graph_1_subtitle, :graph_2_title, :graph_2_subtitle, :image_path, :image_path_2,:image_path_3, :image_path_4, :logo, :data], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @data = data
        @image_name = File.basename(image_path)
        @image_name_2 = File.basename(image_path_2)
        @image_name_3 = File.basename(image_path_3)
        @image_name_4 = File.basename(image_path_4)
        @logo_name = File.basename(logo)
      end

      def save(extract_path, index)
        puts image_path
        puts image_path_2
        puts image_path_3
        puts image_path_4
        puts logo
        puts '------------------->'
        copy_media(extract_path, image_path)
        copy_media(extract_path, image_path_2)
        copy_media(extract_path, image_path_3)
        copy_media(extract_path, image_path_4)
        copy_media(extract_path, logo)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def save_rel_xml(extract_path, index)
        render_view('dashboard_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('dashboard.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end