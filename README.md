# Refractored Stock Analysis

## Overview of Project
 
### Purpose & Background
 The purpose of this project was to analyize specific stocks based on years and report a summary in order to make investment decisions. We were also tasked with formatting the results to be easier to read, as well as reporting the run time on the created macro. Essentially we refractored code to run more efficiently and also gave results for a wider array of data.
## Results

### Analysis
Based on the returned results we can conclude that not only does the code run more efficiently with less clutter but it also provides easy to read data in order to better understand the chosen stocks. It also provided us with a larger variety of data by allowing us to choose specific years as well as select stock options. Specific code like "RowCount = Cells(Rows.Count, "A").End(xlUp).Row" allowed us to easily obtain the rows that needed to be analyzed and plug them in. Our cell formatting along with timers and a simple messageBox made our results very easy to read.

https://github.com/william11329/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true
https://github.com/william11329/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true

## Summary

### Advantages and disadvantages of refactoring code in general
The obvious advantages to code refractoring are that our code becomes easier to read after being organized. with more organization and less clutter peers are able to read code easier and perhaps find errors easier. This can also lead to an easier understanding of your goal and allow for faster programming. A disadvantage to refractoring could be time if the code is to complex or possibly preventing the code from working if done incorrectly.
### Advantages and disadvantages of the original and refactored VBA script
After timing my origional code vs the refractored code for the same 2018 dataset the new refractored code runs faster. It seems to me that from this specfic refractoring on these datasets the only downside was the time it actually took to change the code. Otherwise the code runs faster, more efficiently and allows input for different data sheets.


