require 'rubygems'
require 'watir-webdriver'

#  
#  This class store all system config and variable
#  @author : NGA1HC
# 

class GlobalVariables
  
  $timeOut = "60000"
  $shorttimeOut = 15
  $longtimeOut = 90
  $reportTimeOut = 45000 

  $Excel_filename = 'FTC_AutomotiveAftermarket_AA_EN_V1.0.xls'
  $execExcelFile = ""
  $execResultFile = ""
  $img = ""
  $FirefoxProfile = ""
  $ScenarioExecution = "Scenario"
  $ScenarioHeader = "ScenarioName"
  $TestCaseExecution = "Test Cases"
  $InputData = "InputData"
  $Results = "Result"

  $result = Array.new(2) 
  $failreason = ""
  $actualreason = ""
  $imgScreenShoot = ""

  $var1 = 0
  $var2 = 0
  $var3 = ""
  $var4 = ""

  $HubAddress = ""
  $num_CrossBrowser = 4
  $CrossBrowser = 0
  $Firefox10 = "C:/Program Files (x86)/Mozilla Firefox/firefox.exe"
  $Firefox36 = "C:/Program Files (x86)/Mozilla Firefox3.6/firefox.exe"
  
end
