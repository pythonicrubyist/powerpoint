require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  class Powerpoint::Presentation

    attr_reader :pptx_path, :extract_path

    def initialize
      @slide_count = 0
      @extract_path =  "extracts/extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}"
      FileUtils.copy_entry TEMPLATE_PATH, @extract_path
    end

    def add_intro title, subtitile=nil
      @slide_count += 1
      Powerpoint::Slide::Intro.new @extract_path, title, subtitile
    end

    def add_textual_slide title, content=[]
      @slide_count += 1
      Powerpoint::Slide::Textual.new @extract_path, title, content, @slide_count
      Powerpoint::Slide::Relationship.new @extract_path, @slide_count
      setablish_relationships
    end

    def add_pictorial_slide title, image_path
      @slide_count += 1
      Powerpoint::Slide::Pictorial.new @extract_path, title, image_path, @slide_count
      setablish_relationships
    end    
    
    def setablish_relationships
      if @slide_count > 2
        Powerpoint::ContentType.new @extract_path, @slide_count
        Powerpoint::Relationship.new @extract_path, @slide_count
        Powerpoint::Meta.new @extract_path, @slide_count
      end
    end

    def save path
      @pptx_path = path
      File.delete(path) if File.exist?(path)
      Powerpoint.compress_pptx @extract_path, @pptx_path
      FileUtils.rm_rf(@extract_path)
      path
    end
  end
end