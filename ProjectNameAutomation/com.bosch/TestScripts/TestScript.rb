require 'watir-webdriver'
require_relative '../Library/ApplicationLibrary/AppFunc.rb'

# 
#  This is all test methods for the project
#  @author : NGA1HC  
# 

class TestScript
  
  def initialize
    @@app = AppFunc.new
  end
  
  def openApp
    begin
      @@app.enterURLToBrowser()
    rescue Exception => e  
    end
  end
  
  def gotoMenu(menu)
    begin
      @@app.clickMatchValue(HomeVar.groupMenu, menu.to_s, "Menu")
    rescue Exception => e  
    end
  end
  
  def homeVerifyRelatedLink(relatedLink)
    begin
      @@app.clickMatchValue(HomeVar.lnkRelated, relatedLink.to_s, "Related Link")
    rescue Exception => e  
    end
  end
  
  def viewProduct(image)
    begin
      @@app.clickMatchValue(ProductVar.imgProduct, image, "Spark Plugs")
    rescue Exception => e  
    end
  end
  
end
