# Stock_Analysis_VBA_Challenge.xlsm
Challenge #2 analyzes stocks using VBA to automate
## Overview of Project

### Purpose and Background

- The purpose of this assignment was to look through an Excel data set 
  of daily stock returns for 2017 and 2018, use VBA code to automate a 
  stock analysis of total daily volume and yearly returns. In this
  assignment, we needed to become familiar with VBA subroutines, syntax,
  formating, and methods. We were also introduced to functions such as
  "For" loops, and and conditional "If" statements. As always, the
  ultimate purpose is to learn and get familiar with coding.

## Results

### 2017 vs 2018 Performance

- Based on the data from "Steve", 2017 stocks in general, outperformed,
  2018 stocks. "TERP" was the only stock in the 2017 portfolio to have a 
  loss. "ENPH" and "RUN" were the only stocks in 2018 to have a gain.
  (See "VBA_runtimes_png" folder)


### Original vs Refactored Execution Time

- The refactored code performed approximately 1 second faster than the 
  original. The reason is that in the original code, the for loop we 
  created to cycle through all rows, cycles through each time we report 
  an outcome. The refactored code makes use of arrays so that it only
  needs to cycle through the data one time to report all outcomes.
  (See "VBA_runtimes_png" folder)
  
## Summary

1) The advantages of refactoring code is that the code becomes less
   computationally expensive. The new code is more efficient and as a 
   result, runs faster. The disadvantage is that the code is a bit more 
   difficult for a novice to write.

2) The original scrip was easier to write but less efficient. Also, I
   discovered that the original script had to be re-formated each time
   it was ran. For example, if I ran the analysis on 2017 and pressed the
   format button, the 2017 formatting remained even after I ran a 2018
   analysis. Since the formatting was embedded in the refactored code,
   the all stocks analysis would auto format.
