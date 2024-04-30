// template filling and exporting PDFs for the 2021 MCO report cards ;



#include "C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\SFY2023-2024\\MCO Report Card - 2024\\Program\\6. Graphics\\Program\\TableToPDF.jsx" ;

var     filesDirectory = "C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\SFY2023-2024\\MCO Report Card - 2024\\Program\\6. Graphics\\",
        templatePath = filesDirectory + "Data\\InDesign_Templates\\Developing_files\\",
        dataPath = filesDirectory + "Output\\"
        ;

// function calls -- edit this part with your template and data files
MailMergeAndStar(templatePath + "MCO Report Cards 2024 - STAR Child.indd", 
                            dataPath + "STAR Child\\MCO Report Cards - SC - bySDA-Final.xls") ;

MailMergeAndStar(templatePath + "MCO Report Cards 2024 - STAR Child-ES.indd", 
                            dataPath + "STAR Child\\MCO Report Cards - SC - bySDA-Final.xls") ;

MailMergeAndStar(templatePath + "MCO Report Cards 2024 - STAR Adult.indd", 
                            dataPath + "STAR Adult\\MCO Report Cards - SA - bySDA-Final.xls") ;

MailMergeAndStar(templatePath + "MCO Report Cards 2024 - STAR Adult-ES.indd", 
                            dataPath + "STAR Adult\\MCO Report Cards - SA - bySDA-Final.xls") ;

MailMergeAndStar(templatePath + "MCO Report Cards 2024 - STAR+PLUS.indd", 
                            dataPath + "STAR+PLUS\\MCO Report Cards - SP - bySDA-Final .xls") ;

MailMergeAndStar(templatePath + "MCO Report Cards 2024 - STAR+PLUS-ES.indd", 
                            dataPath + "STAR+PLUS\\MCO Report Cards - SP - bySDA-Final .xls") ;

MailMergeAndStar(templatePath + "MCO Report Cards 2024 - STAR Kids.indd", 
                            dataPath + "STAR Kids\\MCO Report Cards - SK - bySDA-Final .xls") ;

MailMergeAndStar(templatePath + "MCO Report Cards 2024 - STAR Kids-ES.indd", 
                            dataPath + "STAR Kids\\MCO Report Cards - SK - bySDA-Final .xls") ;

