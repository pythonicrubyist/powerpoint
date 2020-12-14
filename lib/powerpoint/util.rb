module Powerpoint
  module Util

    def pixle_to_pt(px)
      px * 12700
    end

    def render_view(template_name, path, variables = {})
      view_contents = read_template(template_name)
      renderer = ERB.new(view_contents)
      b = merge_variables(binding, variables)
      data = renderer.result(b)

      File.open(path, 'w') { |f| f << data }
    end

    def read_template(filename)
      File.read("#{Powerpoint::VIEW_PATH}/#{filename}")
    end

    def require_arguments(required_arguements, arguements)
      puts required_arguements.inspect
      puts arguements.inspect
      raise ArgumentError unless required_arguements.all? {|required_key| arguements.keys.include? required_key}
    end

    def copy_media(extract_path, image_path)
      image_name = File.basename(image_path)
      dest_path = "#{extract_path}/ppt/media/#{image_name}"
      FileUtils.copy_file(image_path, dest_path) unless File.exist?(dest_path)
    end

    def merge_variables(b, variables)
      return b if variables.empty?
      variables.each do |k,v|
        b.local_variable_set(k, v)
      end
      b
    end
  end
end
