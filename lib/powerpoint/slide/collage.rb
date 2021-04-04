require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Collage
      include Powerpoint::Util

      attr_reader :title, :subtitle,:page_number,:logo, :task_icon, :image_information, :images

      def initialize(options={})
        require_arguments [:title, :subtitle,:page_number,:logo, :task_icon, :image_information, :images], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @images = images
        @image_y_scale = image_information[1]
        @image_x_scale = image_information[0]
        @image_x_offset = image_information[2]
        @image_y_offset = image_information[3]
      end

      def save(extract_path, index)
        @images.each do |image|
          copy_media(extract_path, image)
        end
        
        copy_media(extract_path, logo)
        copy_media(extract_path, task_icon)

        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end


      def save_rel_xml(extract_path, index)
        render_view('collage_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('collage.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end