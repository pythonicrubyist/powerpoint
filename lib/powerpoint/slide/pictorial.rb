require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Pictorial
      include Powerpoint::Util

    	attr_reader :image_name, :title, :coords, :image_path

    	def initialize(options={})
				require_arguments [:presentation, :title, :image_path], options
      	options.each {|k, v| instance_variable_set("@#{k}", v)}
        @coords = default_coords unless @coords.any?
      	@image_name = File.basename(@image_path)
      end

      def save(index)
        copy_media_to_extract_path
        save_rel_xml(index)
        save_slide_xml(index)
      end

      def file_type
        File.extname(image_name).gsub('.', '')
      end

    	private

      def default_coords
        slide_width = pixle_to_pt(720); slide_height = pixle_to_pt(540); default_width = pixle_to_pt(550)
        return {} unless dimensions = FastImage.size(image_path)
        image_width, image_height = dimensions.map {|d| pixle_to_pt(d)}
        new_width = default_width < image_width ? default_width : image_width
        ratio = new_width / image_width.to_f
        new_height = (image_height.to_f * ratio).round
        {x: (slide_width / 2) - (new_width/2), y: pixle_to_pt(120), cx: new_width, cy: new_height}
      end

      def copy_media_to_extract_path
        FileUtils.copy_file(@image_path, "#{extract_path}/ppt/media/#{@image_name}")
      end

      def save_rel_xml index
        File.open("#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", 'w'){ |f| f << render_view('pictorial_rel.xml.erb') }
      end

      def save_slide_xml index
        File.open("#{extract_path}/ppt/slides/slide#{index}.xml", 'w'){ |f| f << render_view('pictorial_slide.xml.erb') }
      end

    end
  end
end