require 'win32ole'
require "watir-webdriver"
require 'test/unit'
require_relative '../../Library/StandardLibrary/Environment.rb'
require_relative '../../Library/StandardLibrary/GlobalVariables.rb'

# 
#  This class is Driver Script: drive our Automation Framework by Unit
#  @author : NGA1HC
# 

class AutomationFWTestCase < Environment
  
  @@colSceName = 1
  @@colSceStatus = 2
  @@colTestCaseNumber = 3
  @@colRunNumber = 4
  @@colTestCaseStatus = 5
  @@colTestCaseName = 6
  @@colAutomationStepName = 7
  @@colStepNumber = 9
  @@colStepDescription = 10
  @@colExpectedResults = 11
  @@colObtainedResults = 12
  @@colAutomationStepStatus = 13
  @@colScreenShot = 13
  @@colStartTime = 15
  @@colEndTime = 16
  @@isSkipStep = false
  @@isSkipTestcase = false
  @@iScenario = 0
  @@iTestCase = 0
  @@browsers = Array.new
  @@testData = Array.new
  @@listScenario = Array.new(){Array.new}
  @@listTestcase = Array.new(){Array.new}
  @@listStep = Array.new(){Array.new}
  
  # 
  #  The main function to drive our Automation Framework 
  #  @author : NGA1HC
  #  @raise [Exception, String]
  # 
  
  
  def test_driverScript
    begin
      if($CrossBrowser < @@browsers.size())
        $execResultFile = @@utils.createNewExecFile(@@browsers.at($CrossBrowser.to_i).to_s)
        @@driver.setup(@@browsers.at($CrossBrowser.to_i))
        
        # Get list Scenario from Scenario sheet
        @@listScenario = @@utils.getScenario($execExcelFile)
        index = 1
        for i in 1..@@listScenario.size() do
          @@isSkipTestcase = false
          index = index + 1
          @@utils.writeTestResult($execResultFile, @@colSceName, index, @@listScenario[i - 1][0].to_s)
          @@iScenario = index
          
        # Get list test case from Scenario
          @@listTestcase = @@utils.getListTestcase($execExcelFile,
                @@listScenario[i - 1][0].to_s)
          @@utils.writeTestResult($execResultFile, @@colStartTime, index, @@driver.DateTimestamp())
          for j in 1..@@listTestcase.size() do
            @@isSkipStep = false
            runtime = Integer(@@listTestcase[j - 1][1].to_s)
            temp = Array.new
            temp = @@listTestcase[j - 1][2].to_s.gsub("[","").gsub("]","").gsub(" ","").split(",")
            setData = Array.new
            if(temp[0].to_s != "0")
              for z in 0..temp.size() do
                setData << temp[z]
              end
            end
            if(runtime < setData.size())
              runtime = setData.size()
            end
            for m in 1..runtime do
              @@isSkipStep = false
              index = index + 1
              @@utils.writeTestResult($execResultFile, @@colTestCaseNumber, index, @@listTestcase[j - 1][0].to_s)
              @@utils.writeTestResult($execResultFile, @@colTestCaseName, index, @@listTestcase[j - 1][3].to_s)
              @@iTestCase = index
              @@utils.writeTestResult($execResultFile, @@colRunNumber, index, Integer(m))       
              # Get list step name from Test case
              @@listStep = @@utils.getStepName($execResultFile, @@listTestcase[j - 1][0].to_s)
              if(@@listStep.size() == 0)
                @@utils.writeTestResult($execResultFile, @@colTestCaseStatus, index, "Test cases don't have step")
              end
              rowScenario = @@iScenario
              colScenario = 1
              rowTestCase = @@iTestCase
              colTestCase = 3
              @@utils.writeTestResult($execResultFile, @@colStartTime, index, @@driver.DateTimestamp())
              for k in 1..@@listStep.size() do
                index = index + 1
                @@utils.writeTestResult($execResultFile,
                      @@colAutomationStepName, index, @@listStep[k - 1][0].to_s)
                @@utils.writeTestResult($execResultFile,
                      @@colStepNumber, index, @@listStep[k - 1][1].to_s)
                @@utils.writeTestResult($execResultFile,
                      @@colStepDescription, index, @@listStep[k - 1][2].to_s)
                @@utils.writeTestResult($execResultFile,
                      @@colExpectedResults, index, @@listStep[k - 1][3].to_s)
                @@utils.writeTestResult($execResultFile,
                      @@colStartTime, index, @@driver.DateTimestamp())
                exec = false
                if(@@isSkipStep == false && @@isSkipTestcase == false)
                  if (setData.size() == 0)
                    @@testData = @@utils.getTestData($execResultFile, @@listTestcase[j - 1][0].to_s, @@listStep[k - 1][0].to_s, m)
                  elsif (setData.size() > 0) 
                    @@testData = @@utils.getTestData($execResultFile, @@listTestcase[j - 1][0].to_s, @@listStep[k - 1][0].to_s, Integer(setData.at(m - 1)))
                  end
                  exec = @@utils.execStep("TestScript", @@listStep[k - 1][0].to_s, @@testData)
                  @@utils.writeTestResult($execResultFile, @@colEndTime, index, @@driver.DateTimestamp())
                  if(exec == false)
                    @@utils.writeTestResult($execResultFile,
                          @@colAutomationStepStatus, index, "Invalid Step name")
                     break
                  end
                else
                   $result[0] = "skip"
                end
                excel = WIN32OLE::connect('excel.Application')
                workbook = excel.Workbooks.Open($execResultFile)
                sheettemp = workbook.Worksheets($Results)
                status = sheettemp.Cells(rowTestCase, colTestCase + 2).Value
                sce_status = sheettemp.Cells(rowScenario, colScenario + 1).Value
                if($result[0].to_s == "true")
                  @@utils.writeTestResult($execResultFile, @@colObtainedResults, index, $result[1].to_s)
                  @@utils.writeTestResult($execResultFile, @@colAutomationStepStatus, index, "PASS")
                  if(k == @@listStep.size())
                    if(status != "FAIL")
                      @@utils.writeTestResult($execResultFile, colTestCase + 2, rowTestCase, "PASS")
                    end
                    if(j == @@listTestcase.size())
                      if(sce_status != "FAIL")
                        @@utils.writeTestResult($execResultFile, colScenario + 1, rowScenario, "PASS")
                      end
                    end
                  end
                elsif ($result[0].to_s == "false")
                  @@utils.writeTestResult($execResultFile, @@colObtainedResults, index, $result[1].to_s)
                  @@utils.writeTestResult($execResultFile, @@colAutomationStepStatus, index, "FAIL")
                  @@utils.writeComment($execResultFile, @@colScreenShot, index, $img, "Link")
                  @@utils.writeTestResult($execResultFile, colTestCase + 2, rowTestCase, "FAIL")
                  @@utils.writeTestResult($execResultFile, colScenario + 1, rowScenario, "FAIL")
                  @@isSkipStep = true
                elsif ($result[0].to_s == "skip")
                  @@utils.writeTestResult($execResultFile, @@colAutomationStepStatus, index, "SKIP")
                  @@utils.writeTestResult($execResultFile, @@colScreenShot, index, "SKIP because previous step is FAIL")
                  if(status != "FAIL")
                    @@utils.writeTestResult($execResultFile, colTestCase + 2, rowTestCase, "SKIP")
                  end
                end          
              end  
              index = index + 1
              @@utils.spaceBetweenTestcase($execResultFile, index)
            end 
          end 
        end 
        @@utils.fillColor($execResultFile) 
        @@utils.drawChart($execResultFile)
        if(@@utils.setResultToJenkins($execResultFile))
          assert(false, "Build fail because have some testcase fail in project") 
        end           
      end
    rescue Exception => e
      puts( e.class )
      puts( e )
    end
  end

  #
  #  Setup is run before every test
  #  @author : NGA1HC
  # 
  
  def setup
    path = File.expand_path('../../../..') 
    path = path.gsub("/","\\")
    $execResultFile = path + "\\Results\\"
    $imgScreenShoot = path + "\\Results\\IMG\\"
    $execExcelFile = path + "\\DataInput\\" + $Excel_filename
    $FirefoxProfile = "FirefoxProfile"
#     Get list browsers
    @@browsers = Utils.new.getBrowser($execExcelFile)
    
  end
  
  # teardown is run after every test
  def teardown
    
  end
  
end


