Remove sheet from excel workbook                                                                                
                                                                                                                
If you have IML you can call R from SAS or                                                                      
use my utl_submit_r64 macro                                                                                     
                                                                                                                
github                                                                                                          
https://tinyurl.com/yyjrf56g                                                                                    
https://github.com/rogerjdeangelis/utl-remove-sheet-from-excel-workbook                                         
                                                                                                                
macros                                                                                                          
https://tinyurl.com/y9nfugth                                                                                    
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories                      
                                                                                                                
other excel repos                                                                                               
https://tinyurl.com/y3p2pqcs                                                                                    
https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=            
                                                                                                                
SAS Forum                                                                                                       
https://tinyurl.com/y5tobyro                                                                                    
https://communities.sas.com/t5/SAS-Programming/How-to-delete-sheet-from-excel-using-sas/m-p/581329              
                                                                                                                
*_                   _                                                                                          
(_)_ __  _ __  _   _| |_                                                                                        
| | '_ \| '_ \| | | | __|                                                                                       
| | | | | |_) | |_| | |_                                                                                        
|_|_| |_| .__/ \__,_|\__|                                                                                       
        |_|                                                                                                     
;                                                                                                               
                                                                                                                
* create workbook with two sheets 'CLASS_M' and sheet 'CLASS_F'                                                 
%utlfkil(d:/xls/class.xlsx);                                                                                    
                                                                                                                
libname xel "d:/xls/class.xlsx";                                                                                
data xel.class_m xel.class_f;                                                                                   
 set sashelp.class;                                                                                             
 select (sex);                                                                                                  
   when ('M') output xel.class_m;                                                                               
   when ('F') output xel.class_f;                                                                               
end; /* no need for otherwise */                                                                                
run;quit;                                                                                                       
libname xel clear;                                                                                              
                                                                                                                
                                                                                                                
   WORKBOOK d:/xls/class_m.xlsx with sheet class                                                                
                                                                                                                
   d:/xls/class_m.xlsx                                                                                          
                                                                                                                
      +----------------------------------------------------------------+                                        
      |     A      |    B       |     C      |    D       |    E       |                                        
      +----------------------------------------------------------------+                                        
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |                                        
      +------------+------------+------------+------------+------------+                                        
   2  | ALFRED     |    M       |    14      |    69      |  112.5     |                                        
      +------------+------------+------------+------------+------------+                                        
   3  | AL         |    M       |    13      |    58      |  101.5     |                                        
      +------------+------------+------------+------------+------------+                                        
       ...                                                                                                      
      +------------+------------+------------+------------+------------+                                        
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |                                        
      +------------+------------+------------+------------+------------+                                        
                                                                                                                
   [CLASS_M]                                                                                                    
                                                                                                                
                                                                                                                
   d:/xls/class_m.xlsx                                                                                          
                                                                                                                
      +----------------------------------------------------------------+                                        
      |     A      |    B       |     C      |    D       |    E       |                                        
      +----------------------------------------------------------------+                                        
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |                                        
      +------------+------------+------------+------------+------------+                                        
   2  | ALICE      |    F       |    14      |    69      |  112.5     |                                        
      +------------+------------+------------+------------+------------+                                        
   3  | MARY       |    F       |    13      |    58      |  101.5     |                                        
      +------------+------------+------------+------------+------------+                                        
       ...                                                                                                      
      +------------+------------+------------+------------+------------+                                        
   20 | WILMA      |    F       |    15      |   66.5     |  112       |                                        
      +------------+------------+------------+------------+------------+                                        
                                                                                                                
   [CLASS_F]                                                                                                    
                                                                                                                
                                                                                                                
*            _               _                                                                                  
  ___  _   _| |_ _ __  _   _| |_                                                                                
 / _ \| | | | __| '_ \| | | | __|                                                                               
| (_) | |_| | |_| |_) | |_| | |_                                                                                
 \___/ \__,_|\__| .__/ \__,_|\__|                                                                               
                |_|                                                                                             
;                                                                                                               
                                                                                                                
* Just sheet 'class_f'                                                                                          
                                                                                                                
                                                                                                                
   d:/xls/class_m.xlsx                                                                                          
                                                                                                                
      +----------------------------------------------------------------+                                        
      |     A      |    B       |     C      |    D       |    E       |                                        
      +----------------------------------------------------------------+                                        
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |                                        
      +------------+------------+------------+------------+------------+                                        
   2  | ALICE      |    F       |    14      |    69      |  112.5     |                                        
      +------------+------------+------------+------------+------------+                                        
   3  | MARY       |    F       |    13      |    58      |  101.5     |                                        
      +------------+------------+------------+------------+------------+                                        
       ...                                                                                                      
      +------------+------------+------------+------------+------------+                                        
   20 | WILMA      |    F       |    15      |   66.5     |  112       |                                        
      +------------+------------+------------+------------+------------+                                        
                                                                                                                
   [CLASS_F]                                                                                                    
                                                                                                                
*          _       _   _                                                                                        
 ___  ___ | |_   _| |_(_) ___  _ __                                                                             
/ __|/ _ \| | | | | __| |/ _ \| '_ \                                                                            
\__ \ (_) | | |_| | |_| | (_) | | | |                                                                           
|___/\___/|_|\__,_|\__|_|\___/|_| |_|                                                                           
                                                                                                                
;                                                                                                               
                                                                                                                
                                                                                                                
%utl_submit_r64('                                                                                               
library(XLConnect);                                                                                             
wb <- loadWorkbook("d:/xls/class.xlsx");                                                                        
removeSheet(wb, sheet = "class_m");                                                                             
saveWorkbook(wb);                                                                                               
');                                                                                                             
                                                                                                                
                                                                                                                
                                                                                                                
