require 'zip/filesystem'
require 'fileutils'
require 'tmpdir'

module Powerpoint
  class Presentation
    include Powerpoint::Util

    attr_reader :slides

    def initialize
      @slides = []
    end

    def add_intro(title, subtitile = nil)
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::Intro}[0]
      slide = Powerpoint::Slide::Intro.new(presentation: self, title: title, subtitile: subtitile)
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide 
      else
        @slides.insert 0, slide
      end
    end

    def add_extended_intro(title, image_path, image_path_2,subtitle = nil,subtitle_2 = nil, coords = {})
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::ExtendedIntro}[0]
      slide = Powerpoint::Slide::ExtendedIntro.new(presentation: self, title: title, subtitle: subtitle, image_path: image_path, image_path_2: image_path_2, subtitle_2: subtitle_2, coords: coords)
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide 
      else
        @slides.insert 0, slide
      end
    end

    def add_textual_slide(title, content = [])
      @slides << Powerpoint::Slide::Textual.new(presentation: self, title: title, content: content)
    end

    def add_pictorial_slide(title, image_path, coords = {})
      @slides << Powerpoint::Slide::Pictorial.new(presentation: self, title: title, image_path: image_path, coords: coords)
    end

    def add_text_picture_slide(title, image_path, content = [])
      @slides << Powerpoint::Slide::TextPicSplit.new(presentation: self, title: title, image_path: image_path, content: content)
    end

    def add_picture_description_slide(title, image_path, content = [])
      @slides << Powerpoint::Slide::DescriptionPic.new(presentation: self, title: title, image_path: image_path, content: content)
    end

    def add_multiple_image_slide(title, subtitle = nil, page_number, logo, task_icon ,images)
      @slides << Powerpoint::Slide::MultipleImage.new(presentation: self, title: title, subtitle: subtitle, page_number: page_number, task_icon: task_icon, logo: logo, images: images)
    end

    def add_concept_slide(title, subtitle = nil, page_number, logo, task_icon, image_information ,images)
      @slides << Powerpoint::Slide::Concept.new(presentation: self, title: title, subtitle: subtitle, page_number: page_number, logo: logo, task_icon: task_icon, image_information: image_information, images: images)
    end

    def add_image_slide(title, subtitle = nil,images)
      @slides << Powerpoint::Slide::Image.new(presentation: self, title: title, subtitle: subtitle)
    end

    def add_dashboard_slide(title, subtitle = nil, page_number, graph_1_title, graph_1_subtitle, graph_2_title, image_1, image_2, image_3, image_4, logo, data)
      @slides << Powerpoint::Slide::Dashboard.new(presentation: self, title: title, subtitle: subtitle, page_number: page_number, graph_1_title: graph_1_title, graph_1_subtitle: graph_1_subtitle, graph_2_title: graph_2_title, image_path: image_1, image_path_2: image_2, image_path_3: image_3, image_path_4: image_4, logo: logo, data: data)
    end

    def add_dashboard_user_slide(title, subtitle = nil, page_number, logo, images)
      @slides << Powerpoint::Slide::DashboardUser.new(presentation: self, title: title, subtitle: subtitle,page_number: page_number, logo: logo, images: images)
    end

    def save(path)
      Dir.mktmpdir do |dir|
        extract_path = "#{dir}/extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}"

        # Copy template to temp path
        FileUtils.copy_entry(TEMPLATE_PATH, extract_path)

        # Remove keep files
        Dir.glob("#{extract_path}/**/.keep").each do |keep_file|
          FileUtils.rm_rf(keep_file)
        end

        # Render/save generic stuff
        render_view('content_type.xml.erb', "#{extract_path}/[Content_Types].xml")
        render_view('presentation.xml.rel.erb', "#{extract_path}/ppt/_rels/presentation.xml.rels")
        render_view('presentation.xml.erb', "#{extract_path}/ppt/presentation.xml")
        render_view('app.xml.erb', "#{extract_path}/docProps/app.xml")

        # Save slides
        slides.each_with_index do |slide, index|
          slide.save(extract_path, index + 1)
        end

        # Create .pptx file
        File.delete(path) if File.exist?(path)
        Powerpoint.compress_pptx(extract_path, path)
      end

      path
    end

    def file_types
      
    end
  end
end
