require "powerpoint/version"
require 'powerpoint/util'
require 'powerpoint/slide/intro'
require 'powerpoint/slide/extended_intro'
require 'powerpoint/slide/multiple_image'
require 'powerpoint/slide/image'
require 'powerpoint/slide/textual'
require 'powerpoint/slide/pictorial'
require 'powerpoint/slide/text_picture_split'
require 'powerpoint/slide/picture_description'
require 'powerpoint/slide/dashboard_users'
require 'powerpoint/slide/dashboard'
require 'powerpoint/compression'
require 'powerpoint/presentation'

module Powerpoint
  ROOT_PATH = File.expand_path("../..", __FILE__)
  TEMPLATE_PATH = "#{ROOT_PATH}/template/"
  VIEW_PATH = "#{ROOT_PATH}/lib/powerpoint/views/"
end
