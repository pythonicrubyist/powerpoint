# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'powerpoint/version'

Gem::Specification.new do |spec|
  spec.name          = "powerpoint"
  spec.version       = Powerpoint::VERSION
  spec.authors       = ["pythonicrubyist"]
  spec.email         = ["pythonicrubyist@gmail.com"]
  spec.description   = %q{A Ruby gem that can create a PowerPoint presentation.}
  spec.summary       = %q{powerpoint is a Ruby gem that can create a PowerPoint presentation based on a standard pptx template.}
  spec.homepage      = "https://github.com/pythonicrubyist/powerpoint"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.required_ruby_version = '>= 1.9.2'

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
  spec.add_development_dependency 'rspec', '~> 2.13.0'
  spec.add_development_dependency 'pry'


  spec.add_dependency 'rubyzip', '~> 1'
  spec.add_dependency 'fastimage', '~> 1'
end
