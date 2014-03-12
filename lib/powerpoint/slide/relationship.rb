require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  module Slide
    class Powerpoint::Slide::Relationship
      def initialize extract_path, slide_number
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/></Relationships>'
        path = "#{extract_path}/ppt/slides/_rels/slide#{slide_number}.xml.rels"
        File.open(path, 'w'){ |f| f << xml }
      end
    end
  end
end