
module Powerpoint
  module Utils
    class Table
      def initialize matrix, settings = {}
        raise ArgumentError.new("no rows") if matrix.first.nil?
        raise ArgumentError.new("no cols") if matrix.first.first.nil?
        
        @matrix   = matrix
        @ncols    = matrix.first.length
        @width    = (settings[:width]    or "1645920")
        @style    = (settings[:style]    or "5C22544A-7EE6-4342-B048-85BDC9FD1C3A")
        @firstrow = (settings[:firstrow] or "1")
        @bandrow  = (settings[:bandrow]  or "1")
        
      end

      def width=(width)
        @width = width
      end

      def style=(style)
        @style = style
      end

      def firstrow=(firstrow)
        @firstrow = firstrow
      end

      def bandrow=(bandrow)
        @bandrow = bandrow
      end

      def to_xml
		%{
			<a:tbl>
				<a:tblPr firstRow="#{@firstrow}" bandRow="">
					<a:tableStyleId>{#{@style}}</a:tableStyleId>
				</a:tblPr>
				<a:tblGrid>
					#{
						"<a:gridCol w='#{@width}' />"*@ncols
					}
				</a:tblGrid>
				#{
					@matrix.map do |row|
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
				}
			</a:tbl>
		}
      end
    end
  end
end
