# stock-analysis
Space to hold information from Module 2 VBA
## Background
In this project, we are helping Steve and his parents understand green energy stocks from 2017-2018. 
We are trying to measure the yearly return and total daily volume. 
Through VBA we are automating the process of looking through the 17/18' data. 
For this module challenge we have both the 'green stocks' and 'VBa_Challenge' files. The green stocks file is our inital output, whereas the second file uses refractoring to make the timing more efficient (moving away from nested for loops). 
Steve's parents are invested in DAQO (DQ stock). We want to compare this to our tickers and see if it is worth investing. 
For outputs, we have the ticker, total daily volume, and the return as a percentage with condotional formatting 

## Results
<img width="399" alt="Stocks Compared" src="https://user-images.githubusercontent.com/102266450/163507235-13fddd10-b1c7-4c36-8344-f2671a3471b1.png">

By looking at both the stock outputs for 2017 and 2018; we see the following: 

Returns
* 2017 outpeformed 2018 for green stocks in positive (green marked) returns. 
* 2018 had a majority of green stocks in the negative percentage of returns (red marked).
* The DQ ticker peformed worse year over year, dropping from 200% positive returns in 17' to -63% in 18'
* Recommendation: Hold and see its value at 2019, but place a sell off option if Steve's parents inital investment drops below 20%. We may suggest to put their funds in a less volatile stock. 

Daily Volume
* If the DQ ticker showed high daily trading, we could rationalize holding on the investment. 
* Once again, due to its 2018 performance, we found that its daily trading volume had no positive impact on the -63% drop for the year. This stock in particular is very risky. 
* Taking a closer look at 2018 (a down year for most green investments); we did identify "ENPH" (+81%) and "RUN" (+84%) as viable options for investment. In fact, going back to their 2017 performance, (ENPH - +129%, RUN +5.5%) makes us question if Steve's parents should consider selling on the DQ stock. 
* However, we must also consider that the 2017 volume for each display over 200,000,000 daily trades vs the DQ daily volume of 35,000,000. This supports our theory that the daily volume could translate to higher performing stocks. 

## Comparing AllStockAnalysis vs Refactored Version
Although both have the same run output (different timing); we do notice the following: 

Ticker
* We created a variable as a string to hold 12 elements since we have 12 tickers to track. The tickerIndex function allows us to index and return values from the table
![loop ref](https://user-images.githubusercontent.com/102266450/163507954-e76e31b5-9d99-4437-99c2-ca1d37488d92.PNG)

This is a much more efficient way of running through our tickers, whereas previosuly our code was as follows: 
![loop non ref](https://user-images.githubusercontent.com/102266450/163508019-a706c207-516c-4373-9d79-3b725eeda093.PNG)

Run Time
![VBA Non Ref 2017](https://user-images.githubusercontent.com/102266450/163507369-e317225a-a468-4090-8997-a2c609770cf8.png)
![VBA Non Ref 2018](https://user-images.githubusercontent.com/102266450/163507370-9a2298da-4269-4c98-a190-b4bb862d83a2.png)


![VBA_Challenge_2017](https://user-images.githubusercontent.com/102266450/163507379-d066c705-0100-4816-b47c-65edcbfaaaa7.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/102266450/163507381-0204f136-1f6e-47ea-9735-4de6719fdffb.png)

We saw that a 140% difference between the run time of .0625 (2018) refactored versus the non-refactored run-time of .4218 (2018). Although the task of automating our code without nested loops was challenging, we can clearly see that speed matters---especially when making financial decisions. 

## Advantages and Disadvantages to refactoring
Personally, refactoring the VBA script was quite difficult. Initial frustrations came from my inability to undestand indexes and how to remove the nested loop. 
However, when the speed presents a large difference, it is hard to argue against the refractored version. Moving forward, I would state that frustrations in building code fall less on your ability to do somethin 'quicky'; and more on one's ability to spend time on the logic and look to your resources to build out the functions step by step. I found myself starting and stopping quite often, which helped pace myself against the want to 'complete' the code. 
I will state that all codes/scripts built, including the DQ analysis have its own purpose. Thinking more broadly, helping Steven from a large funnel approach works if he wants to help clients (his parents in this case) see how all energy stocks are peforming in a given year. But as I think about this project more closely, I can see the value in only displaying one ticker output at a time. If we narrow our focus, we may have better insights than presenting all information by year and 11 tickers. In another approach, we may resuse the DQ analysis script and replace its ticker to filter for another stock.  

