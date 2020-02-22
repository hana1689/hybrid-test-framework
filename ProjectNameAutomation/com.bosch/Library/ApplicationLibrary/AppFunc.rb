require 'rubygems'
require 'watir-webdriver'
require_relative '../StandardLibrary/GlobalVariables.rb'
require_relative '../../Library/ApplicationLibrary/HomeVar.rb'
require_relative '../../Framework/ExcelFunctions/Utils.rb'
require_relative '../../Library/StandardLibrary/StandardFunc.rb'

# 
#  This is all test methods for the project
#  @author : NGA1HC  
# 

class AppFunc
  
  def initialize
    @@utils = Utils.new
    @@driver = StandardFunc.new
  end

  def checkPageTemplate
    begin
      @@driver.verifyIsElementPresent(HomeVar.AATitle, "AA Title")
      @@driver.verifyIsElementPresent(HomeVar.AALogo, "AA Logo")
      @@driver.verifyIsElementPresent(HomeVar.AAMenu, "AA Menu")
      @@driver.verifyIsElementPresent(HomeVar.AAContent, "AAContent")
      @@driver.verifyIsElementPresent(HomeVar.AAFooter, "AAFooter")
    rescue Exception => e
      @@utils.setFalseResult("VERIFY Page Template - FAIL") 
      raise Exception.new
    end
  end
  
  def loginApp(xpathUser, xpathPassword, xpathButton, xpathLoginSuccess, dataUser, dataPassword)
    begin
      @@driver.type_text(xpathUser, dataUser, "User Login")
      @@driver.input_password(xpathPassword, dataPassword, "Password Login")
      @@driver.click_button(xpathButton, "Button Login")
      @@driver.verifyIsElementPresent(xpathLoginSuccess, "Login success")
      @@utils.setTrueResult("VERIFY login application -- PASS")
    rescue Exception => e 
      raise Exception.new
    end
  end
  
  def logoutApp(xpathLogout)
    begin
      @@driver.click_button(xpathLogout, "Logout")
      checkPageTemplate()
      @@utils.setTrueResult("VERIFY logout application -- PASS")
    rescue Exception => e
      raise Exception.new  
    end
  end
  
  def clickMatchValue(loc, value, element)
    begin
      href = @@driver.getAttributeHref(loc, value)
      @@driver.click_link("text = #{value}")
      id = @@driver.get_url()
      if(href.to_s != id.to_s)
        @@utils.setFalseResult("VERIFY link another page -- FAIL REDIRECT TO ANOTHER PAGE")
        raise Exception.new
      else
        @@utils.setTrueResult("VERIFY " + element + " "  + value + " -- PASS")
      end
    rescue Exception => e
      @@driver.recoveryMode()
    end
  end

  def enterURLToBrowser
    begin
      globalVar = Array.new
      globalVar = @@utils.getGlobalVars($execResultFile)
      @@driver.openURL(globalVar.at(0))
      checkPageTemplate()
      @@utils.setTrueResult("VERIFY open Application " + globalVar.at(0) + " -- PASS")
    rescue Exception => e
      raise Exception.new  
    end
  end
  
  
end