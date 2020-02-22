require 'win32ole'
require 'watir-webdriver'
require 'rubygems'
require_relative '../../Library/StandardLibrary/GlobalVariables.rb'
require_relative '../../Library/StandardLibrary/StandardFunc.rb'
require_relative '../../TestScripts/TestScript.rb'

# 
#  This class store all function working with Excel file
#  @author : NGA1HC
# 

class Utils
  
  @@colSce=['A','B','C','D','E','F','G','H','I','J','K','L','M']
  # @@excel = WIN32OLE::new("Excel.Application")
  
  def initialize
    begin
      @@excel = WIN32OLE::connect('excel.Application')
    rescue Exception => e 
       @@excel = WIN32OLE::new("Excel.Application")
    end
  end
 
  def lastIndexSheet(excelpath, excel_sheet, array)
    i=6
    workbook = @@excel.Workbooks.Open(excelpath)
    sheet =  workbook.WorkSheets(excel_sheet) 
    for column in array do 
      while (sheet.Range(column+"#{i}").value!=nil or sheet.Range(column+"#{i+1}").value!=nil)
        i=i+1
      end 
    end 
    workbook.Close
    return i
  end
   
  def findrecell(excelpath, excel_sheet, keywork)
    begin
    index = lastIndexSheet(excelpath,excel_sheet,["A","B","C","D"])
    address=[]
    workbook = @@excel.Workbooks.Open(excelpath)
    sheet =  workbook.WorkSheets(excel_sheet) 
    for column in @@colSce do 
      for i in 1..index do
      data=sheet.Range(column +"#{i}").value
        if data!=nil and  data.strip !=nil 
          data=sheet.Range(column +"#{i}").value.strip
        else 
          data =sheet.Range(column +"#{i}").value
        end     
        if data==keywork
          address[0]=column
          address[1]=i          
        break
        end        
      end
      if (address!=nil)
         break
      end
    end 
    workbook.Close
    return address
    rescue Exception => e
      puts( e.class )
      puts( e )
      puts( "Can't found #{keywork} in sheet")  
    return nil 
    end
  end

  #  
  #  @author : NGA1HC
  #  @param [excelpath, String] 
  #  @return [listScenario, String[][]] includes scenario name with flag yes, test case name
  #  @raise [Exception, String] 
  # 

  def getScenario(excelpath)
    begin     
      numScenario = 0
      numSceFlag = 0
      index = -1
      workbook = @@excel.workbooks.Open(excelpath)
      sheet = workbook.WorkSheets($ScenarioExecution)     
      rowSce=3 
      i= 0
      
    #       Get size for Scenario object
      begin
        while (sheet.Range(@@colSce[0]+"#{rowSce}").value !=nil)
            rowSce=rowSce+1    
            i=i+1 
            numScenario = i-1
        end
      rescue Exception => e
        numScenario = i-1
      end
     # Get all scenario with Flag is Yes  
     rowSce=3
          for j in 1.. numScenario do 
           sceFlag= sheet.Range(@@colSce[1]+"#{rowSce+j}").value     
           if sceFlag=="Yes"
           numSceFlag = numSceFlag+1
           end    
          end   
       rowSce=3    
      # Set size for scenario
      listScenario=Array.new(numSceFlag){Array.new(3) }
         for j in 1..numScenario do 
         sceName=sheet.Range(@@colSce[0]+"#{rowSce+j}").value     
         sceFlag=sheet.Range(@@colSce[1]+"#{rowSce+j}").value
         sceTestCase=sheet.Range(@@colSce[2]+"#{rowSce+j}").value
        
         if (sceName !="" and sceFlag=="Yes") 
          index = index+1
          listScenario[index][0] = sceName
          listScenario[index][1] = sceFlag
          listScenario[index][2] = sceTestCase
        elsif sceName==""
          break
        end
      end    
      workbook.Close           
     return listScenario  
  rescue Exception => e   
    puts( e.class )
    puts( e )
    puts( "Can't get Scenario List" )  
    return nil
  end
end

  #  
  #  @author : NGA1HC
  #  @param [excelpath, String] 
  #  @param [scenarioName, String] 
  #  @return [ret_obj, Array] includes test case number with flag yes, 
  #    execution times and data (if have), test case name 
  #  @raise [Exception, String] 
  # 

  def getListTestcase(excelpath, scenarioName)
    begin
      listTestCase= []
      listTestCase1= []
      times =""
      dataset = Array.new
      workbook = @@excel.Workbooks.Open(excelpath)
      sheet1 = workbook.WorkSheets($ScenarioExecution)
      sheet2 =  workbook.WorkSheets($TestCaseExecution) 
       
       i= 4   
               while (sheet1.Range(@@colSce[0]+"#{i}").value !=nil)
                 if (sheet1.Range(@@colSce[0]+"#{i}").value== scenarioName)
                     sheet1.Range("C"+"#{i}").value.split(/\s*,\s*/).each do|c|
                     listTestCase<< c    
                     end       
                   break
                  else
                      i=i+1
                  end
               end
    
  #   Get list Test case with Flag is Yes
        m= listTestCase.size()
        for i in 0..m do
          for row in 6.. sheet2.UsedRange.Rows.Count 
               if (sheet2.Range(@@colSce[0]+"#{row}").value== listTestCase[i] and sheet2.Range(@@colSce[6]+"#{row}").value=="Yes")
                  listTestCase1 << listTestCase[i]
                  break
               end
           end
         end
  
        testcasesize= listTestCase1.size()
      
        #  Set size for list Test Case
       ret_obj = Array.new(listTestCase1.size()){Array.new(4) }
       
        # Get all TestCase and put return object
        for k in 0 .. testcasesize-1  
           for i in 6.. sheet2.UsedRange.Rows.Count
            if (sheet2.Range(@@colSce[0]+"#{i}").value == listTestCase1[k])
               if (sheet2.Range(@@colSce[7]+"#{i}").value != nil)
                 runtime = sheet2.Range(@@colSce[7]+"#{i}").value
                 runtime = runtime.to_s.gsub(" ","").split(/,#*/)
                 times = runtime.at(0).to_i
           
                 ret_obj[k][0] = listTestCase1[k]
                 ret_obj[k][1] = times
                 ret_obj[k][3] = sheet2.Range(@@colSce[4]+"#{i}").value
                 if (runtime.size() == 1)
                   ret_obj[k][2] = 0
                 else
                   for i in 1..runtime.size() do
                     dataset << runtime[i]
                   end
                   ret_obj[k][2] = dataset
                 end
              end    
            end              
           end            
          end
          workbook.Close    
          return ret_obj
      rescue Exception => e
      puts( e.class )
      puts( e )
      puts( "Can't get list Test Case for Scenario : #{scenarioName}")  
      return nil
     end  
  end

  #  
  #  @author : NGA1HC
  #  @param [excelpath, String] 
  #  @return [browser, String[]] 
  #  @raise [Exception, String] 
  #
   
  def getBrowser(excelpath)
   browser =[] 
   workbook = @@excel.workbooks.Open(excelpath) 
   sheet1= workbook.WorkSheets($ScenarioExecution)
   begin
     for i in ["B","C","D"] do 
        browser<< sheet1.Range("#{i}"+"1").value   
     end
     workbook.Close
     return browser 
    rescue Exception => e
      puts( e.class )
      puts( e )
      puts( "Can't get list browser")  
      return nil
     end   
   end  
 
  #  
  #  @author : NGA1HC
  #  @param [excelpath, String] 
  #  @param [testcaseName, String] 
  #  @return [listStep, Array] includes stepname, step number, description, expected result
  #  @raise [Exception, String] 
  # 
   
  def getStepName(excelpath, testcaseName)
    begin   
     # find position of Step Name
     index2 = lastIndexSheet(excelpath,$TestCaseExecution,["j","k","l","m"])
      # sheet1 = @workbook.WorkSheets($ScenarioExecution)
     workbook = @@excel.Workbooks.Open(excelpath)
     sheet2 = workbook.WorkSheets($TestCaseExecution) 
      
     for i in 6 ..index2 do
       # puts i
        if(sheet2.Range(@@colSce[0]+"#{i}").value==testcaseName)
          startStepName= i
          break
        end
      end
  start = startStepName+1
     for i in start ..index2 do
       if (sheet2.Range(@@colSce[0]+"#{i}").value!=nil or sheet2.Range(@@colSce[9]+"#{i}").value==nil)
       endStepname = i - 1
       break
       end
     end
     stepsize = endStepname-startStepName+1
     # puts stepsize
     listStep= Array.new(stepsize){Array.new(4)}
     j=0
     for i in startStepName ..endStepname do 
        stepname= sheet2.Range(@@colSce[9]+"#{i}").value
        steps= sheet2.Range(@@colSce[10]+"#{i}").value
        decription= sheet2.Range(@@colSce[11]+"#{i}").value
         result= sheet2.Range(@@colSce[12]+"#{i}").value
      # puts i
         listStep[j][0]= stepname
         listStep[j][1]= steps
         listStep[j][2]= decription
         listStep[j][3]= result
       j=j+1
          
      end
   
  workbook.Close
  return listStep
   rescue Exception => e
      puts( e.class )
      puts( e )
      puts( "Can't get list Test data for test case Name : #{testcaseName}")  
      return nil 
   end
  end
  
  #  
  #  @author : NGA1HC
  #  @param [excelpath, String] 
  #  @param [testcaseName, String]
  #  @param [stepname, String] 
  #  @param [dataset, int]
  #  @return [data, String[]] 
  #  @raise [Exception, String] 
  # 

  def getTestData(excelpath, testcaseName, stepname, dataset)
    cellTestCase=[]
    cellTestCase = findrecell(excelpath,$InputData,testcaseName)
   
    startTestCase= cellTestCase[1]
    endtestcase=0
  begin
    workbook = @@excel.Workbooks.Open(excelpath)
    sheet1 = workbook.WorkSheets($ScenarioExecution) 
    sheet3 =  workbook.WorkSheets($InputData)
    i = startTestCase+1
    while (sheet3.Range(@@colSce[0] +"#{i}").value!=nil or sheet3.Range(@@colSce[2]+"#{i}").value!=nil) 
      if(sheet3.Range(@@colSce[0] +"#{i}").value!=nil and sheet3.Range(@@colSce[2]+"#{i}").value!=nil)
      endTestCase=i
       break
      end
      i=i+1
    end
    # puts startTestCase,endTestCase
    # find position of step name 
    for i in startTestCase ..endTestCase do
      if sheet3.Range(@@colSce[0]+"#{i}").value==stepname
            startStepname =i 
        break
      end
    end
    endStepName = 0
    for i in startStepname ..endTestCase do
      if (i == endTestCase) 
          endStepName = endTestCase
          break
      end
      if(sheet3.Range(@@colSce[0]+"#{i+1}").value!=nil)     
          endStepName = i
          break
      end   
       
     end
    # puts startStepname, endStepName
    data=[]
    start= startStepname+1
      for i in start..endStepName do
      value =dataset + 1      
      data<< sheet3.Range(@@colSce[value]+"#{i}").value
      end 
        workbook.Close
        return data
     rescue Exception => e
      puts( e.class )
      puts( e )
      puts( "Can't get list Test data for testcase Name : #{testcaseName} and step name #{stepname}")  
      return nil 
     end
   end

  #  
  #  @author : NGA1HC
  #  @param [excelpath, String] 
  #  @return [var, Array] includes browser, username, password
  #  @raise [Exception, String] 
  # 
 
   def getGlobalVars(excelpath)
     begin
       workbook = @@excel.workbooks.Open(excelpath)
       sheet3 =  workbook.WorkSheets($InputData)
       var = Array.new
       for i in 0..3 do
         var[i] = sheet3.Range(@@colSce[1]+"#{i+3}").Value
       end
       workbook.Close
       return var
     rescue Exception => e
       puts "Can't get list global vars"
     end
   end
   
  #  
  #  @author : NGA1HC
  #  @param [classname, String] 
  #  @param [stepname, String] 
  #  @param [params, String[]] 
  #  @return [isExec, boolean] 
  #  @raise [Exception, String] 
  #

  def execStep(classname, stepname, params)
    begin
  	isExec = false
  	methods = Array.new  	
	  cls = Object.const_get(classname)
    obj = cls.new
    methods = cls.instance_methods(false)
    listMethod = Array.new
    for i in 0..methods.size() do
      listMethod << methods[i]
    end
    param = Array.new
    param = params
    if(listMethod.to_s.include?(stepname))
      isExec = true
      if(param == nil)
        obj.send(stepname)
      else
        case param.size()
        when 1
          obj.send(stepname, param[0])
        when 2
          obj.send(stepname, param[0], param[1])
        when 3
          obj.send(stepname, param[0], param[1], param[2])
        end
      end
    end
    return isExec
    rescue Exception => e
      puts( e.class )
      puts( e )
      puts( "Can't found data for Step Name : #{stepname}")  
      return nil 
    end
  end
  
  #  
  #  @author : NGA1HC
  #  @param [browser, String] 
  #  @return [outputFile, String] is the path of excel result file
  #  @raise [Exception, String] 
  #
  
  def createNewExecFile(browser)
    begin
    excel_size = $Excel_filename.size() - 5
    preoutputFile = $Excel_filename[0..excel_size] + '_' + browser + DateTimestamp() + '.xls'
    postoutputFile = preoutputFile.gsub(" ","")
    formattedOutput = postoutputFile.gsub(":","")
    
    outputFile = $execResultFile + formattedOutput
    
    originalFile = @@excel.Workbooks.Open($execExcelFile)
    destinationFile = @@excel.Workbooks.Add
    
    for i in 1..4 do
      originalFile.Worksheets(i).Copy(destinationFile.Worksheets(i))
    end
    
    # @destinationFile.Name = outputFile
    destinationFile.SaveAs(outputFile)
    destinationFile.Close
    $execResultFile = outputFile
    return outputFile
  rescue Exception => e
    puts( e.class )
    puts( e )
    puts( "Can't create new excel file")
    return nil
  end
  end
  
  #  
  #  @author : NGA1HC
  #  @return [date, String] is the current date
  #
  
  def DateTimestamp
    return Time.new.strftime("%Y-%m-%d%H%M%S")
  end
  
  #  
  #  @author : NGA1HC
  #  @param [excelpath, String]
  #  @param [index, int] is the row position
  #  @raise [Exception, String] 
  #
  
  def spaceBetweenTestcase(excelpath, index)
    begin
      workbook = @@excel.Workbooks.Open(excelpath)
      sheet = workbook.WorkSheets($Results)      
      for j in 1..15 do
        if(sheet.Cells(index,j).Value == nil)
          sheet.Cells(index, j).Interior.ColorIndex = 15
        end
      end
      workbook.Save
      workbook.Close     
  rescue Exception => e
    puts( e.class )
    puts( e )
    puts( "Can't write comment")
    return nil
  end
  end
  
  #  
  #  @author : NGA1HC
  #  @param [excelpath, String]
  #  @raise [Exception, String] 
  #
  
  def fillColor(excelpath)
    begin
      workbook = @@excel.Workbooks.Open(excelpath)
      sheet = workbook.WorkSheets($Results)
      
      columns = sheet.UsedRange.Columns.Count
      rows = sheet.UsedRange.Rows.Count
      for i in 1..columns do
        for j in 1..rows do
          if(sheet.Cells(j, i).Value == "PASS")
            sheet.Cells(j,i).Font.ColorIndex = 10
            sheet.Cells(j,i).Font.Bold = true
          elsif (sheet.Cells(j, i).Value == "FAIL")
            sheet.Cells(j,i).Font.ColorIndex = 3
            sheet.Cells(j,i).Font.Bold = true
          elsif (sheet.Cells(j, i).Value == "SKIP")
            sheet.Cells(j,i).Font.ColorIndex = 13
            sheet.Cells(j,i).Font.Bold = true
          end
        end
      end
      workbook.Save
      workbook.Close
  rescue Exception => e
    puts( e.class )
    puts( e )
    puts( "Can't set color ")
    return nil
  end
  end
  
  #  
  #  @author : NGA1HC
  #  @param [excelpath, String]
  #  @param [column, int]
  #  @param [row, int]
  #  @param [link, String] is the hyperlink to screenshot
  #  @param [desc, String] is the description
  #  @raise [Exception, String] 
  #
  
  def writeComment(excelpath, column, row, link, desc)
    begin
      workbook = @@excel.Workbooks.Open(excelpath)
      sheet = workbook.WorkSheets($Results)
      
      hl = sheet.Hyperlinks.Add(sheet.Cells(row,column), "")
      hl.Address = link
      puts link
      hl.TexttoDisplay = desc
      # drag mouse over link and the text displayed
      hl.ScreenTip = "Click to go to this URL"  
      workbook.Save
      workbook.Close      
  rescue Exception => e
    puts( e.class )
    puts( e )
    puts( "Can't write comment")
    return nil
   end
  end
  
  #  
  #  @author : NGA1HC
  #  @param [excelpath, String]
  #  @raise [Exception, String] 
  #
  
  def drawChart(excelpath)
    begin
      pass = 0
      fail = 0
      skip = 0
      
      workbook = @@excel.Workbooks.Open(excelpath)
      sheet = workbook.WorkSheets($Results)
      rows = sheet.UsedRange.Rows.Count
      for i in 1..rows do
        if (sheet.Cells(i, 5).Value == "PASS")
          pass = pass + 1
        elsif (sheet.Cells(i, 5).Value == "FAIL")
          fail = fail + 1
        elsif (sheet.Cells(i, 5).Value == "SKIP")
          skip = skip + 1
        end
      end
      sheet.Cells(8,30).Value = pass
      sheet.Cells(9,30).Value = fail
      sheet.Cells(10,31).Value = skip
      workbook.Save
      workbook.Close      
  rescue Exception => e
    puts( e.class )
    puts( e )
    puts( "Can't draw chart")
    return nil
  end
  end
  
  #  
  #  @author : NGA1HC
  #  @param [excelpath, String]
  #  @param [column, int]
  #  @param [row, int]
  #  @param [content, String] 
  #  @raise [Exception, String] 
  #
  
  def writeTestResult(excelpath, column, row, content)
    begin
      workbook = @@excel.Workbooks.Open(excelpath)
      
      sheet = workbook.WorkSheets($Results)
      
      sheet.Cells(row, column).Value = content
      workbook.Save
      workbook.Close
  rescue Exception => e
    puts( e.class )
    puts( e )
    puts( "Can't write test result")
    return nil
   end
  end
  
  #  
  #  @author : NGA1HC
  #  @param [actual, String] is the actual result
  #
  
  def setTrueResult(actual)
    $result = Array.new(2)
    $result[0] = "true"
    $result[1] = actual
  end
  
  #  
  #  @author : NGA1HC
  #  @param [reason, String] is the fail reason
  #
  
  def setFalseResult(reason)
    $img = StandardFunc.new.captureScreen()
    $result = Array.new(2)
    $result[0] = "false"
    $result[1] = reason
  end
  
  #  
  #  @author : NGA1HC
  #  @param [excelpath, String] 
  #  @raise [Exception, String] 
  #
  
  def setResultToJenkins(excelpath)
    begin
      setResult = false
      workbook = @@excel.Workbooks.Open(excelpath)
      sheet = workbook.WorkSheets($Results)
      
      rows = sheet.UsedRange.Rows.Count
      for i in 1..rows do
        if(sheet.Cells(i,5).Value == "FAIL")
          setResult = true
        end
      end
    return setResult
  rescue Exception => e
    puts( e.class )
    puts( e )
    puts( "Can't set result to jenkins")
    return nil
   end
 end

  
end



           
           
           
