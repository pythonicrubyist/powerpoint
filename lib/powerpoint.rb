require "powerpoint/version"
require 'powerpoint/util'
require 'powerpoint/slide/intro'
require 'powerpoint/slide/extended_intro'
require 'powerpoint/slide/ranking'
require 'powerpoint/slide/collage'
require 'powerpoint/slide/concept'
require 'powerpoint/slide/gallery'
require 'powerpoint/slide/image'
require 'powerpoint/slide/textual'
require 'powerpoint/slide/pictorial'
require 'powerpoint/slide/text_picture_split'
require 'powerpoint/slide/picture_description'
require 'powerpoint/compression'
require 'powerpoint/presentation'

module Powerpoint
  ROOT_PATH = File.expand_path("../..", __FILE__)
  TEMPLATE_PATH = "#{ROOT_PATH}/template/"
  VIEW_PATH = "#{ROOT_PATH}/lib/powerpoint/views/"
end
