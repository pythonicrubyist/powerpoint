require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  module Slide
    class Powerpoint::Slide::Table
      def initialize extract_path, title, table, slide_number
        #table has to have rows
        raise ArgumentError, "Table must have rows" if table.length == 0

        #get number of cols for the gridcol tags
        ncols = table.first.length

        template_path = "#{TEMPLATE_PATH}/ppt/slides/slide3.xml"
        xml = File.read template_path

        #set the title
        xml.gsub!("TITLE_PLACEHOLDER", title)

        #set the gridcol tags
        cols_xml = ""
        ncols.times do |n|
          cols_xml << "<a:gridCol w='1645920' />"
        end
        xml.gsub! "COLS_PLACEHOLDER", cols_xml

        #make the actual rows
        table_xml = table.map do |row|
          %{
			<a:tr h="370840">
				#{
					row.map do |col|
						%{
							<a:tc>
								<a:txBody>
									<a:bodyPr />
									<a:lstStyle />
									<a:p>
										<a:r>
											<a:rPr lang="en-US" dirty="0" smtClean="0" />
											<a:t>#{col}</a:t>
										</a:r>
										<a:endParaRPr lang="en-US" dirty="0" />
									</a:p>
								</a:txBody>
								<a:tcPr />
							</a:tc>
						}
					end.join
				}
			</a:tr>
		}
        end.join
        xml.gsub! "ROWS_PLACEHOLDER", table_xml

        #write the result to the powerpoint
        slide_path = "#{extract_path}/ppt/slides/slide#{slide_number}.xml"
        File.open(slide_path, "w") { |f| f << xml }
      end
    end
  end
end
