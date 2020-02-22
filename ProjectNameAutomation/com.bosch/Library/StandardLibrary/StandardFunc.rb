require 'rubygems'
require 'watir-webdriver'
require_relative '../../Framework/ExcelFunctions/Utils.rb'
require_relative 'GlobalVariables.rb'

# 
#  This is all methods prepare for automation framework
#  @author : NGA1HC
# 

class StandardFunc 
  
  def initialize
    @@utils = Utils.new
  end
  
  def setup(browser)
    begin
      case (browser)
      when 'Internet Explorer'
        $browser = Watir::Browser.new :ie
        browser_name = 'Internet Explorer'
      when 'Firefox 10'
        Selenium::WebDriver::Firefox.path = $Firefox10
        # profile.add_extension 'autoauth-2.1-fx+fn.xpi'
        $browser = Watir::Browser.new :firefox
      when 'Firefox 36'
        Selenium::WebDriver::Firefox.path = $Firefox36
        # profile.add_extension 'autoauth-2.1-fx+fn.xpi'
        $browser = Watir::Browser.new :firefox
      when 'FirefoxProfile'
        profile = Selenium::WebDriver::Firefox::Profile.from_name $FirefoxProfile
        # profile.add_extension 'autoauth-2.1-fx+fn.xpi'
        $browser = Watir::Browser.new :firefox, :profile => profile
      when 'chrome'
        $browser = Watir::Browser.new :chrome
        browser_name = 'Chrome'
      when 'opera'
        $browser = Watir::Browser.new :opera
        browser_name = 'Opera'
      else
          profile = Selenium::WebDriver::Firefox::Profile.from_name $FirefoxProfile
          # profile.add_extension 'autoauth-2.1-fx+fn.xpi'
          $browser = Watir::Browser.new :firefox, :profile => profile
      end
      MaximizeBrowser()
      implicitlyWait(2)
      rescue Exception => e
        puts( e.class )
        puts( e )
        puts( "Couldn't start with browser #{browser}")  
      end
  end
  
  def MaximizeBrowser
    begin
      $browser.driver.manage.window.maximize
      @@utils.setTrueResult("Maximize browser is executed")
    rescue Exception => e
      puts( e.class )
      puts( e )
    end
  end
  
  def implicitlyWait(time)
    begin
      $browser.driver.manage.timeouts.implicit_wait = time
    rescue Exception => e
      puts( e.class )
      puts( e )
    end
  end
  
  def DateTimestamp
    return Time.new.strftime("%Y-%m-%d%H%M%S")
  end
   
  # Get current URL

  def get_url
    begin
      return $browser.url
    rescue Exception => e
      puts( e.class )
      puts( e )
    end
  end
         
  # Go to specific URL in already-opened browser
  def openURL(url)
    begin
      $browser.goto(url)
      @@utils.setTrueResult("Open \"" + url + "\" is executed")
    rescue Exception => e
      puts( e.class )
      puts( e )
    end
  end
#     
  # Go back in browsing history
  def go_back
    begin
      $browser.back
    rescue Exception => e
      @@utils.setFalseResult("Couldn't go back in Browser")
      puts( e.class )
      puts( e )
    end
  end
    
  # Go forward in browsing history
  def go_forward
    begin
      $browser.forward
    rescue Exception => e
      @@utils.setFalseResult("Couldn't go foward in Browser")
      puts( e.class )
      puts( e )
    end
  end
      
  # Close current browser
  def close_browser
    begin
      $browser.close
    rescue Exception => e
      puts( e.class )
      puts( e )
      puts("Couldn't close browser")
    end
  end
  
  # (Non-keyword) Small utility for trimming whitespace from both sides of a string
  #
  # @param [String] s string to trim both sides of
  # @return [Strgin] the same string without whitespace on the left or right sides
  #
  def trim_sides(s)
    s = s.lstrip
    s = s.rstrip
    return s
  end
  
  def parse_location(loc)
    loc = trim_sides(loc)
    
    if loc[0..3].downcase == 'css='
      return {:css => loc[4..loc.length]}
    elsif loc[0..5].downcase == 'xpath='
      return {:xpath => loc[6..loc.length]}
    elsif !loc.include? '='
      return {:id => loc}
    else
      # Comma-separated attributes
      attr_list = loc.split(',')
      attrs = {}
      attr_list.each do |a|
        attr_kv = a.split('=')
        # Need to turn strings in format "/regex-here/" into actual regexes
        if attr_kv[1].start_with?('/')
          attr_kv[1] = Regexp.new(Regexp.quote(attr_kv[1].gsub('/', '')))
        end
        attr_kv[1] = self.trim_sides(attr_kv[1]) unless attr_kv[1].is_a? Regexp
        attrs[self.trim_sides(attr_kv[0])] = attr_kv[1]
      end
      # Watir expects symbols for keys
      attrs = Hash[attrs.map { |k, v| [k.to_sym, v] }]
      if attrs.key? :id
        attrs.delete_if {|k, v| k != :id}
      end
      return attrs
    end
  end
  
  # Input text into a text field
  # @param [String] loc attribute/value pairs that match an HTML element
  # @param [String] text the text to type into the field
  def input_text(loc, text)
    begin
      $browser.text_field(parse_location(loc)).set text
    rescue Exception => e
      @@utils.setFalseResult("Couldn't type text on text field" + loc)
      puts( e.class )
      puts( e )
    end
  end
     
  def input_textarea(loc, text)
    begin
      $browser.textarea(parse_location(loc)).set text
    rescue Exception => e
      @@utils.setFalseResult("Couldn't type text on textarea" + loc)
      puts( e.class )
      puts( e )
    end
  end

  def input_password(loc, password, element)
    begin
      $browser.text_field(parse_location(loc)).set password
    rescue Exception => e
      @@utils.setFalseResult("Couldn't type password on" + element)
      puts( e.class )
      puts( e )
    end
  end
  
  # Get value of given text field
  def get_textfield_value(loc)
    begin
      return $browser.text_field(parse_location(loc)).value
    rescue Exception => e
      @@utils.setFalseResult(loc + "don't have attribute value")
      puts( e.class )
      puts( e )
    end
  end

  # Click a button
  def click_button(loc, element)
    begin
      $browser.button(parse_location(loc)).click
    rescue Exception => e
      @@utils.setFalseResult("Couldn't click button on" + element)
      puts( e.class )
      puts( e )
    end
  end
     
  # Check a checkbox
  def select_checkbox(loc)
    begin
      $browser.checkbox(parse_location(loc)).set
    rescue Exception => e
      @@utils.setFalseResult("Couldn't select checkbox on" + loc)
      puts( e.class )
      puts( e )
    end
  end
  
  # Get an element's text
  def get_element_text(loc)
    begin
      $browser.element(parse_location(loc)).text
    rescue Exception => e
      @@utils.setFalseResult(loc + "don't have attribute value")
      puts( e.class )
      puts( e )
    end
  end
  
  # Click any HTML element
  def click_element(loc)
    begin
      $browser.element(parse_location(loc)).click
    rescue Exception => e
      @@utils.setFalseResult("Couldn't click on" + loc)
      puts( e.class )
      puts( e )
    end
  end

  # Type text into the given text field (alternate for "Type Text")
  def type_text(loc, text, element)
    begin
      $browser.element(parse_location(loc)).send_keys(text)
    rescue Exception => e
      @@utils.setFalseResult("Couldn't type text on" + element)
      puts( e.class )
      puts( e )
    end
  end
  
  # Verify that an element is visible
  def element_should_be_present(loc)
    begin
      return $browser.element(parse_location(loc)).visible?
    rescue Exception => e
      return false
    end
  end
  
  # Select a radio button
  def select_radio_button(loc)
    begin
      $browser.radio(parse_location(loc)).set
    rescue Exception => e
      @@utils.setFalseResult("Couldn't select radio button on" + loc)
      puts( e.class )
      puts( e )
    end
  end
  
  # Insert path of a file into a file-upload field
  def choose_file(loc, path)
    begin
      $browser.file_field(parse_location(loc)).set path
    rescue Exception => e
      @@utils.setFalseResult("Couldn't select path file on" + loc)
      puts( e.class )
      puts( e )
    end
  end
  
  # Get the file path present in a file-upload field
  def get_filefield_path(loc)
    begin
      $browser.file_field(parse_location(loc)).value
    rescue Exception => e
      @@utils.setFalseResult(loc + "don't have attribute value")
      puts( e.class )
      puts( e )
    end
  end

  # Get text from the items of a list
  #
  # The second argument can be "true" or "false"; true for ordered list, false (default) for unordered.
  #
  # @param [String] loc attribute/value pairs that match an HTML element
  # @param [String] ordered true is ol, false is ul; strings for the booleans
  # @param [String] sep the separator that should be used to separate the list items
  # 
  def get_list_items(loc, ordered, sep = ";;")
    begin
      ordered = ordered.downcase unless ordered.nil?
      if ordered.nil?
        list = $browser.ul(parse_location(loc))
      elsif ordered == 'true'
        list = $browser.ol(parse_location(loc))
      elsif ordered == 'false'
        list = $browser.ul(parse_location(loc))
      end       
      items = []
      list.lis.each do |item|
        items << item.text
      end       
      items.join(sep)
    rescue Exception => e
      @@utils.setFalseResult("Couldn't get list items on" + loc)
      puts( e.class )
      puts( e )
    end
  end
  
  # Click a single table cell
  #
  # @param [String] loc attribute/value pairs that match an HTML element
  # @param [String] row with the top-most row as 1, the number of the row in question
  # @param [String] col with the left-most row as 1, the number of the column in question
  # 
  def click_table_cell(loc, row, col)
    begin
      row = row.to_i - 1
      col = col.to_i - 1
      $browser.table(parse_location(loc))[row][col].click
    rescue Exception => e
      @@utils.setFalseResult("Couldn't click on table " + loc)
      puts( e.class )
      puts( e )
    end
  end
  
  # Count how many row within an entire HTML table
  def table_row(loc)
    begin
      return $browser.table(parse_location(loc)).rows.size()
    rescue Exception => e
      @@utils.setFalseResult(loc + "don't have row")
      puts( e.class )
      puts( e )
      return nil
    end
  end

  # Get page title
  def get_title
    begin
      return $browser.title
    rescue Exception => e
      @@utils.setFalseResult("Couldn't get title")
      puts( e.class )
      puts( e )
    end
  end

  # Get the text of all elements that match a given XPath query
  #
  # @param [String] xpath the xpath query to use for searching
  # @param [String] sep the separator that will be used in printing out the results
  # @return [String] a string of the text of all the matching elements, separated by sep
  #
  def get_all_elements_by_xpath(xpath, sep = ';;')
    begin
      matches = []
      $browser.elements_by_xpath(xpath).each do |element|
        matches << element.text
      end
      return matches.join(sep)
    rescue Exception => e
      @@utils.setFalseResult("Couldn't find list element" + xpath)
      puts( e.class )
      puts( e )
      return nil
    end
  end

  # Click a link
  def click_link(loc)
    begin
      $browser.link(parse_location(loc)).click
    rescue Exception => e
      @@utils.setFalseResult("Couldn't click on link " + loc)
      puts( e.class )
      puts( e )
    end
  end

  # Click an image
  def click_image(loc)
    begin
      $browser.image(parse_location(loc)).click
    rescue Exception => e
      @@utils.setFalseResult("Couldn't click on image " + loc)
      puts( e.class )
      puts( e )
    end
  end
  
  # Capture screenshot if test case failed
  def captureScreen
    begin
      outputFile = ""
      preoutputFile = DateTimestamp() + "IMG.png"
      postoutputFile = preoutputFile.gsub(" ", "")
      formattedOutput = postoutputFile.gsub(":", "")
      
      outputFile = $imgScreenShoot + formattedOutput
      $browser.driver.save_screenshot(outputFile)
      return outputFile
    rescue Exception => e
      @@utils.setFalseResult("Couldn't capture screenshot")
      puts( e.class )
      puts( e )
      return nil
    end
  end
  
  def getAttributeHref(loc, value)
    begin
      href = $browser.div(parse_location(loc)).a(parse_location("text = #{value}")).attribute_value "href"
      return href.to_s
    rescue Exception => e
      raise Exception.new  
    end
  end
  
  def click_href()
    begin
      
    rescue Exception => e
        
    end
  end
  
  def verifyIsEqual(expected, actual, reason)
    begin
      if(expected != actual)
        @@utils.setFalseResult(reason)
      end
    rescue Exception => e
      puts( e.class )
      puts( e ) 
    end
  end
  
  def verifyIsFalse(condidion, reason)
    begin
      if(condition == true)
        @@utils.setFalseResult(reason)
      end
    rescue Exception => e
      puts( e.class )
      puts( e )   
    end
  end
  
  def verifyIsTrue(condition, reason)
    begin
      if(condition == false)
        @@utils.setFalseResult(reason)
      end
    rescue Exception => e
      puts( e.class )
      puts( e )   
    end
  end    
  
  def verifyIsElementPresent(loc, element)
    begin
      if(element_should_be_present(loc) != true)
        @@utils.setFalseResult("Verify " + element + " is present fail")
        raise Exception.new
      end
    rescue Exception => e 
      raise Exception.new  
    end
  end
  
  def recoveryMode
    begin
      browsers = @@utils.getGlobalVars($execExcelFile)
      $browser.goto(browsers.at(0))
    rescue Exception => e
      raise Exception.new  
    end
  end
    
    
    
end


