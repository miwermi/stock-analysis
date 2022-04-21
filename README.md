                                                                                           Michelle Werner (4/17/2022)
# Stock Analysis: VBA Code Refactoring & Measuring Performance 
---

## Overview of Project

Intially this project was about client Steve, who was evaluating green energy stocks for his parents. The initial analysis of the DQ stock they had chosen indicated that it had lost value and wasn't the best investment. Following that discovery, Steve asked for help to design a program that could look at multiple stocks simultaneously and compare changes in their stock values. 

![Steve's Stock](resources/SteveStockAnalysis.png)

Pictured: Steve's Stock Analysis 

### Purpose

When writing a program to compare multiple stocks and their total value increases and decreases, and thinking of how the program might be need to "scale" up even further in the future, a few things seemed relevant: visual clarity, variety and timing. Adding buttons and visual formatting to the program made it more easily understood at a glance, and the program was much improved with the added ability to access different spreadsheets from multiple stock years. The final refactoring of the program involves invesitgating timing.

As the data comparison demand increases, saving time on the computation becomes useful - and interesting. For this final re-write of the program, we are going to see if we can test the speed of two different code methods used to find  and total the begining and ending price comparisons to see which is more efficient. This type of adjustment to our original code is referred to as "refactoring". Our findings and the value of refactoring this VBA code will be discussed further in the Results section below.

For more on refactoring visit: https://www.bmc.com/blogs/code-refactoring-explained/
---
## Results

Below are six screenshots of our program results, the first two show the results from the original program, the next two are from the first attempt at refactoring in which one array was added, and the last two are the FINAL three-array refactored program images:


![Initial timing, 2018 data](resources/M2_stockanalysis_2018.png)

Figure 1: Initial timing, 2018 Green Stock Analysis 

![Initial timing, 2017 data](resources/M2_stockanalysis_2017.png)

Figure 2: Initial timing, 2017 Green Stock Analysis 

The figures above display the results and timing of our intial program coding. You can see that for the 2018 data, the timing for calculating the results is indicated as 66915.74 seconds and for the 2017 run, the timing is 66889.4 seconds.

### Refactored

Below are the improved timings with our re-factored code. In this case the timings returned are 0.78125 for 2018 and .8046875 for 2017, along with a view of the refactored code.

![One array timing, 2018 data](resources/VBA_Challenge_2018_refactored.png)

Figure 3: Improved timing, 2018 Green Stock Analysis Refactored with one new array

![One array timing, 2017 data](resources/VBA_Challenge_2017_refactored.png)

Figure 4: Improved timing, 2017 Green Stock Analysis Refactored with one new array


![Refactored VBA code](resources/code_refactored.png)

Figure 5: Refactored VBA Challenge Code with one new array

The code above was a little cleaner/more efficient than our orinal code - as evidenced by the timing differences; in this case the timings returned are 0.78125 for 2018 and .8046875 for 2017. Also viewable is the refactored code (Figure 5). 
### FINAL Refactoring with 3 Arrays
Below is the even more improved three-array timings with the FINAL refactored code (Figure 8).

![FINAL three array timing, 2018 data](resources/VBA_Challenge_2018_FINAL.png)

Figure 6: FINAL timing, 2018 Green Stock Analysis Refactored with three arrays

![One array timing, 2017 data](resources/VBA_Challenge_2017_FINAL.png)

Figure 7: FINAL timing, 2017 Green Stock Analysis Refactored with three arrays

![FINAL Refactored VBA code with three arrays](resources/code_FINAL.png)

Figure 8: FINAL Refactored VBA Challenge Code with three arrays

---
## Summary
In general, refactoring code is the practice of cleaning up code. A programmer may write a program with code that is loaded with redundancy (or "dirty") because it is more simple to test and debug or becuase they need to quickly "get er done".  But cleaner, more succinct code is highly valued by anyone who may need to edit it in the future. So it is always a "best practice" to consider refactoring as the cleaner the code, the easier it is to maintain and to add future features to. 

"The act of refactoring – changing tiny pieces of code with no front-end purpose – may seem unimportant when compared to higher priority tasks. But the cumulative effect from such changes is significant and can lead to a better-functioning team and approach to programming."  (BMC blogs,  Stephen Watts &Chrissy Kidd, 2018, https://www.bmc.com/blogs/code-refactoring-explained/)

For our VBA challenge, we have successfully improved our program by adding arrays and simplifying the amount of work the program had to do. Refactoring has allowed us to have cleaner code with more streamlined functionality, but it has also sped up calculaton timing. By adding one more array, "tickerIndex", we were able to have the program loop less, improving efficiency and saving valuable time. Adding three arrays made even more improvements! The FINAL refactored version of our project has an advantage over our previous versions because ultimately, it can handle more calculations in far less time than it took in our original version.

One thing to note though, while the refactored code in our program is definitely faster, the timings that were returned seem a bit off (as the difference between versions was mere seconds and not minutes and/or hours as the timings seem to indicate). 

Figuring that out could be our next challenge!
