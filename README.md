# An Evaluation of Stocks for Future Investment
## Overview of the Project: A summary of Total Daily Volume and Return % for Stocks in 2017-2018 to best determine which ones to make an investment

This project asked me to look at 2 years worth of stock performance for 12 different stocks. It was asked to calculate total Volume and Return percentage to best determine which stocks might be good to invest in. 

## Results

### Stock Performance over the years
In evaluating possible stocks for your parents' future investments, I felt it was important to look at multiple years of stock performance history. It was important to see both years because they give quite different snapshots of return and volume depending on the year. 

#### 2017 Results
2017 was a year with multiple stocks showing high returns of investment (3 with over 100%).  There was only one stock (TERP) that was showing a negative return of investment. 2017 seemed to be a lucrative time for Wall Street investors. I believe your parents were interested in investing in DQ. While 2017 showed them with almost 200% return of investment, 2018 was telling quite a different story. 

![2017_Stock_Return_Summary.png](https://github.com/melaniekelsey/stock-analysis/blob/main/Resources/2017_Stock_Return_Summary.png)

#### 2018 Results
While in 2018, only 2 stocks actually provided a return in investment, despite the fact that there were higher daily volume overall for the evaluated stocks in that particular year. Both ENPH and RUN had a little over 80% return of investment, while all the others were in the red. Looking at DQ in 2018 you will see -62.6% return in investment. 

![2018_Stock_Return_Summary.png](https://github.com/melaniekelsey/stock-analysis/blob/main/Resources/2018_Stock_Return_Summary.png)

#### Conclusion
In looking at that data, I don't feel confident that DQ is the right stock for your parents to invest in. Though 2017 was a great year for DQ, it seems like the stock might be a bit risky with the significant decline in returns in 2018. In looking at other possibilities, ENPH seems to be the winner. With a 2017 Return of 129.5% and a 2018 Return of 81.9%, it seems like a steady stock to invest in. 

### Execution times in Refactored Code

Refactoring the code was the key to a quicker macro performance. In the original script, the first loop goes through all the rows of data, ticker name by ticker name. So it loops through all the rows of data 12 total times for the 12 total ticker names. With over 3000 rows of data per year, repeating the loop 12 times is inefficient. 

In the refactored code, you loop through the rows of stock data first to calculate the volume, starting price and ending price. Following that loop, you go through all the different ticker symbols to get the results of those 3 arrays. It creates a situation with 12 small loops as opposed to 12 large loops and speeds up the code performance. 

#### Original Script 2017 Performance

![VBA_Challenge_2017.png](https://github.com/melaniekelsey/stock-analysis/blob/main/Resources/VBA_Challenge_2017_OrigScript.png)

#### Refactored Script 2017 Performance

![VBA_Challenge_2017.png](https://github.com/melaniekelsey/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

#### Original Script 2018 Performance

![VBA_Challenge_2018.png](https://github.com/melaniekelsey/stock-analysis/blob/main/Resources/VBA_Challenge_2018_OrigScript.png)

#### Refactored Script 2018 Performance

![VBA_Challenge_2018.png](https://github.com/melaniekelsey/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### What are the advantages of refactoring code?

The advantages of refactoring the code are basically speeding up the performance of the query. You save essentially half a second per query with the refactored code. It might seem insignificant, but this is just a small sample of stock data. If the client were to look at a much larger sample of stock data, the refactored code would be worth it in the long run. 

In general, refactoring code should help make queries more efficient and give you a chance to see if there are any bugs in your code. For me personally, it gave me the chance to breakdown the code and make sure I understood what each line of code was doing overall. 

### What are the disadvantages of refactoring code?

The disadvantages of refactoring code are that an analyst could possibly "break" the existing code. It was pretty complex for me to go through the refactoring process. I had to reread and test lots of different iterations of the code to get this one to work. If this had an earlier due date, I might not have been able to truly refactor this code and have the queries still work effectively.

### How do these pros and cons apply to refactoring the original VBA script?

The refactoring of the original script allowed me to be able to make the query more efficient. In the case that the client wants to see more years of data or additional stocks to analyze, I am making the job easier for both myself and the client to quickly access the information that he needs. 

The cons of the refactoring process are that since I am relatively new at this process, it's hard to tell if my new code is working accurately without having to audit and compare all the data. Also, if I didn't have the modules to review and show me examples of the steps, it would have been challenging to get this process done quickly. 
