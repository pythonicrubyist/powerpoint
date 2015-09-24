module Powerpoint
  module Util

    def pixle_to_pt(px)
      px * 12700
    end

    def render_view(filename)
      view_contents = read_template(filename)
      renderer = ERB.new(view_contents)
      renderer.result( binding )
    end

    def read_template(filename)
      File.read(Powerpoint::ROOT_PATH + "/lib/powerpoint/views/#{filename}")
    end

    def require_arguments(required_argements, argements)
      raise ArgumentError unless required_argements.all? {|required_key| argements.keys.include? required_key}
    end

    def extract_path
      @presentation.extract_path
    end
  end
end