require 'watir-webdriver'

class HomeVar
  
    def self.AATitle; @@AATitle = "websiteTitleArea"; end 
    def self.AALogo; @@AALogo = "xpath = /html/body/div/div/div/a/img"; end
    def self.AAMenu; @@AAMenu = "mainNav"; end
    def self.AAContent; @@AAContent = "contentArea"; end
    def self.AAFooter; @@AAFooter = "footer"; end
    def self.groupMenu; @@groupMenu = "class = mainNavWrapper"; end
    def self.lnkRelated; @@lnkRelated = "class = grid12 floatLe"; end

end
