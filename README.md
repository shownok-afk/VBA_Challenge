# VBA_Challenge

# Study of Different Stocks 

### Overview of Project

This study focuses on different stocks daily trade volumes and returns over the period of two years. It will help, investors to compare performance of different stocks between 2017 and 2018. By the end of this project, we will be able to see if there is any relationship between stocks return and daily trade volume. 

### Stock Analysis Results

### 2017 Results

In the year of 2017 almost all the stocks except TERP performed well. Each stock had positive returns. Specifically, DQ, ENPH, FSLR and SEDG had return over 100%. Though we can see trade volume for DQ was low compare other stocks. 

![2017 Table](https://github.com/shownok-afk/VBA_Challenge/blob/main/Resources/2017%20Table.png)

### 2018 Results
In the year of 2018 we see a complete opposite picture for most of the stocks. Only ENPH and RUN did well in terms of return. Both the stocks are remain consistent over the period of analysis time. We can see DQ is not cosistent over 2018, more over DQ has highest negative return in 2018. From this we can say that DQ is the most volatile stock of our analysis set and ENPH and RUN are most steady stocks to invest. 

![2018 Table](https://github.com/shownok-afk/VBA_Challenge/blob/main/Resources/2018%20Table.png)

## Code refactoring result
Original code runtime is 0.81 second for 2017 and 0.78 for 2018. After code refectoring runtime improved significantly  as it took 0.12s to run both 2017 and 2018.

![VBA_Challenge_2017](https://github.com/shownok-afk/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.png) ![VBA_Challenge_2018](https://github.com/shownok-afk/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png) 
##Summary


### Advantage of refactoring codes 

•	It makes code efficient, so runtime decreases significantly.
•	It takes less resouces to get processed.
•	It becomes more user faiendly.
### Disadvantage of refactoring code

•	Codes some time may get complicated to understand.

### Pros and Cons apply to refactoring the original VBA

In original VBA code only one array was declared. As a result output was written each time when a particular ticker calculation was finished. After refactoring of the codes, which include declare of different arrays and use of tickerindex made program little complicated but increased efficiency as entire calculation was done and stored in differnet array. Also stored information was written to different cells using one forloop and declared arrays. Altogether refactoring made VBA code efficient and it takes less time to run now.
