require 'zip/filesystem'
require 'fileutils'
require 'tmpdir'

module Powerpoint
  class Presentation
    include Powerpoint::Util

    attr_reader :pptx_path, :extract_path, :slides

    def initialize
      @slides = []
    end

    def add_intro title, subtitile=nil
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::Intro}[0]
      slide = Powerpoint::Slide::Intro.new(presentation: self, title: title, subtitile: subtitile)
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide 
      else
        @slides.insert 0, slide
      end
    end

    def add_textual_slide title, content=[]
      @slides << Powerpoint::Slide::Textual.new(presentation: self, title: title, content: content)
    end

    def add_pictorial_slide title, image_path, coords={}
      @slides << Powerpoint::Slide::Pictorial.new(presentation: self, title: title, image_path: image_path, coords: coords)
    end

    def save path
      # Copy template to temp path
      @extract_path = File.join(Dir.tmpdir, "extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}")
      FileUtils.copy_entry TEMPLATE_PATH, @extract_path

      # Remove keep files
      FileUtils.rm_rf("#{@extract_path}/ppt/_rels/.keep")
      FileUtils.rm_rf("#{@extract_path}/ppt/media/.keep")
      FileUtils.rm_rf("#{@extract_path}/ppt/slides/_rels/.keep")

      # Render/save generic stuff
      File.open("#{extract_path}/[Content_Types].xml", 'w') { |f| f << render_view('content_type.xml.erb') }
      File.open("#{extract_path}/ppt/_rels/presentation.xml.rels", 'w') { |f| f << render_view('presentation.xml.rel.erb') }
      File.open("#{extract_path}/ppt/presentation.xml", 'w') { |f| f << render_view('presentation.xml.erb') }

      # Save slides
      slides.each_with_index do |slide, index|
        slide.save(index + 1)
      end

      # Create .pptx file
      @pptx_path = path
      File.delete(path) if File.exist?(path)
      Powerpoint.compress_pptx @extract_path, @pptx_path
      FileUtils.rm_rf(@extract_path)
      path
    end

    def file_types
      slides.select {|slide| slide.class == Powerpoint::Slide::Pictorial}.map(&:file_type).uniq
    end
  end
end