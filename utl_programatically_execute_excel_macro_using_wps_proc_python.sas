Programatically execute excel vba macro using wps proc python;

 github
 https://tinyurl.com/y846mz9g
 https://github.com/rogerjdeangelis/utl_programatically_execute_excel_macro_using_wps_proc_python

 see
 https://tinyurl.com/y7lmf9y8
 http://jacobjwalker.effectiveeducation.org/blog/2015/01/24/python-script-to-automate-refreshing-an-excel-spreadsheet/

 For input xlsm workbook see dropbox or github
 https://www.dropbox.com/s/tokxyqz4p292z4h/class_final.xlsm?dl=0

 https://tinyurl.com/yaloqcuo
 https://github.com/rogerjdeangelis/utl_programatically_execute_excel_macro_using_wps_proc_python/blob/master/class_final.xlsm


INPUT
=====

  * you need this works in 32 and 64 bit Win 7
  pip install pypiwin32

  The macro enable workbook contains

   Sub sum_weight()
    Range("F21").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
    Range("F22").Select
   End Sub


   https://www.dropbox.com/s/tokxyqz4p292z4h/class_final.xlsm?dl=0
   which I saved locally as d:/xls/class_final.xlsm

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
   2  | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+
   21 |            |            |            |            |  1900.9    | calculated by vba sum_weight macro
      +------------+------------+------------+------------+------------+

      [CLASS]


PROCESS  ( working code xl.Run("sum_weight") );
===============================================

%utl_submit_wps64('
options set=PYTHONHOME "C:\Progra~1\Python~1.5\";
options set=PYTHONPATH "C:\Progra~1\Python~1.5\lib\";
proc python;
submit;
import win32com.client;
import os;
xl = win32com.client.DispatchEx("Excel.Application");
wb = xl.Workbooks.open(r"d:\xls\class_final.xlsm");
xl.Visible = True;
xl.Run("sum_weight");
wb.Save();
xl.Quit();
endsubmit;
run;quit;
');


WANT
====

WORK.CLAA total obs=20

Obs    OBS    NAME       SEX    AGE    HEIGHT    WEIGHT

  1      0    Alfred      M      14     69.0      112.5
  2      1    Alice       F      13     56.5       84.0
 18     17    Thomas      M      11     57.5       85.0
....
 19     18    William     M      15     66.5      112.0
 20      .                        .       .      1900.5   * calculated by macro


*                _               _       _
 _ __ ___   __ _| | _____     __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

;
   Sub sum_weight()
    Range("F21").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
    Range("F22").Select
   End Sub

   https://www.dropbox.com/s/tokxyqz4p292z4h/class_final.xlsm?dl=0
   which I saved locally as d:/xls/class_final.xlsm

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|

;

%utl_submit_wps64('
options set=PYTHONHOME "C:\Progra~1\Python~1.5\";
options set=PYTHONPATH "C:\Progra~1\Python~1.5\lib\";
proc python;
submit;
import win32com.client;
import os;
xl = win32com.client.DispatchEx("Excel.Application");
wb = xl.Workbooks.open(r"d:\xls\class_final.xlsm");
xl.Visible = True;
xl.Run("sum_weight");
wb.Save();
xl.Quit();
endsubmit;
run;quit;
');

LOG

NOTE: AUTOEXEC processing completed

1         options set=PYTHONHOME "C:\Progra~1\Python~1.5\";
2         options set=PYTHONPATH "C:\Progra~1\Python~1.5\lib\";
3         proc python;
4         submit;
5         import win32com.client
6         import os
7         xl = win32com.client.DispatchEx("Excel.Application")
8         wb = xl.Workbooks.open(r"d:\xls\class_final.xlsm")
9         xl.Visible = True
10        xl.Run("sum_weight")
11        wb.Save()
12        xl.Quit()
13        endsubmit;

NOTE: Submitting statements to Python:


14        run;
NOTE: Procedure python step took :
      real time : 3.541
      cpu time  : 0.000



