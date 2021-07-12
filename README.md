# **Stock Analysis**

## **Overview of Project**
In this project, we created and then refactored a VBA script in order to analyze a handful of green energy stocks and evaluate their performance in 2017 and 2018. When executed, the VBA code measures how actively a stock is traded by finding the total daily volume and also reviews the performance of the stock by calculating the yearly return.

## **Results**
#### Stock Performance
In comparing the yearly returns, it is evident that green energy companies were recipients of big investment in 2017; A few stocks (i.e. DQ, ENPH, FSLR, SEDG) more than doubled in price. In 2018, we can see that the growth tapered off as most stocks began to wane from their gains in the prior year. Below is a summary: 
![Yearly Comparison](https://user-images.githubusercontent.com/86018601/125211578-4dd83c80-e275-11eb-9b96-11195c4e3429.png)

#### Code Performance
In the original VBA script, we used a nested for loop which looped through the data several times. That is, it repeated the same code block once for each ticker. The refactored code eliminated this nested loop and instead allowed us to loop through the data just one time. As a result, the execution times were reduced considerably. 

###### *Original Code:*
![nested_loop](https://user-images.githubusercontent.com/86018601/125215447-8550e400-e289-11eb-9f08-5a69e7772edf.png)

###### *Refactored Code:*
![refactored_loop](https://user-images.githubusercontent.com/86018601/125215520-b7fadc80-e289-11eb-8886-b6ead72647d1.png)

In comparing the time it took to execute the original code versus the refactored code, we can see that the execution time is reduced from 0.762 seconds to 0.176 seconds for 2017:

![2017 Execution Time Comparison](https://user-images.githubusercontent.com/86018601/125213471-639f2f00-e280-11eb-96c9-d398643f7d6e.png)


Similarly, the execution time for the 2018 analysis is reduced from 0.719 seconds to 0.168 seconds:

![2018 Execution Time Comparison](https://user-images.githubusercontent.com/86018601/125213491-669a1f80-e280-11eb-817b-ee0c89561cc2.png)

## **Summary**
One of the key advantages of refactoring code is to make it more efficient, thereby boosting system performance which is of increasing importance when dealing with larger datasets. Refactored code is also better organized, which makes the flow more logical and improves the readability which will be especially beneficial to future users. One disadvantage is that it can be very time consuming and ultimately does not change the functionality. Also, by making too many changes, you run the risk of introducing bugs to your script. 

In this project, the refactored code allowed us to loop through the data just one time, compared to the original code which looped through the data several times. The end result was a considerable reduction in execution time when compared to the original execution time. Ultimately however, the reduction in time was rather undetectable when in use. 

