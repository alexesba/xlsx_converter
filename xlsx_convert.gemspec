Gem::Specification.new do |s|
  s.name        = 'xlsx_convert'
  s.version     = '0.0.0'
  s.date        = '2015-07-17'
  s.description = 'A simple xlsx file converter into csv and json gem'
  s.authors     = ['Alejandro Espinoza']
  s.email       = 'alexesba@gmail.com'
  s.files       = ['lib/xlsx_convert.rb']
  s.homepage    = 'http://rubygems.org/gems/hola'
  s.license     = 'MIT'
  s.summary     = 'xlsx converter'
  s.add_runtime_dependency 'simple_xlsx_reader', '~> 1.0', '>= 1.0.2'
  s.add_runtime_dependency 'json', '~> 1.8', '>= 1.8.3'
end
