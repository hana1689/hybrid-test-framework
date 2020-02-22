require 'watir-webdriver'
require 'win32ole'
require 'rubygems'
require 'test/unit'
require_relative 'StandardFunc.rb'
require_relative '../../Framework/ExcelFunctions/Utils.rb'

# 
#  This class will init Standard Library instance and Util instance
#  @author: NGA1HC
#  

class Environment < Test::Unit::TestCase
  
  @@utils = Utils.new
  @@driver = StandardFunc.new

  
end
