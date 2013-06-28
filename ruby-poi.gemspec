# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'poi/version'

Gem::Specification.new do |spec|
  spec.name          = "ruby-poi"
  spec.version       = Ruby::Poi::VERSION
  spec.authors       = ["Ægir Örn Símonarson"]
  spec.email         = ["agirorn@gmail.com"]
  spec.description   = %q{Ruby bridge for Apache POI}
  spec.summary       = %q{Making Office Document manipulations just a little less painful.}
  spec.homepage      = "https://github.com/agirorn/ruby-poi"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_dependency 'rjb'

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
end
