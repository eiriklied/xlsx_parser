# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'xlsx_parser/version'

Gem::Specification.new do |gem|
  gem.name          = "xlsx_parser"
  gem.version       = XlsxParser::VERSION
  gem.authors       = ["Eirik Lied"]
  gem.email         = ["eiriklied@gmail.com"]
  gem.description   = %q{A parser for Xlsx documents (excel) with a low memory footprint so its good for huge xlsx files}
  gem.summary       = gem.description
  gem.homepage      = "https://github.com/eiriklied/xlsx_parser"

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ["lib"]

  gem.add_development_dependency "rspec"
  gem.add_runtime_dependency "nokogiri"
  gem.add_runtime_dependency "rubyzip", ">= 1.0.0"
end
