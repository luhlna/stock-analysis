# Stock Analysis

## Overview
An Stock analysis has been run in order to help Steve and his parents to evaluate investing in renewable energy companies. They wanted to invest in DaQo New Energy Corporation based on a nostalgic reason, but with a set of data from 2017 and 2018 we helped them made the best decision. After evaluating DQ, Steve wanted to extend the analysis to other companies from 2017 and so. 

## Analysis
First we evaluated the daily volume, which measures how the stocks are traded every day. Then we calculate the performance of the stock over a year as the percentage difference in price from the beginning to the end of the year.


To automate and allow Steve evaluate under this parameters any set of data we use the power of Macros with VBA. When running such large sets we wanted to give the user an estimated amount of time that the Macros will take to run, therefore we added a Timer to show it:

<img width="272" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/90527212/135798857-0cf12845-6e45-4c12-9657-fd03949e06c8.png"> <img width="273" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/90527212/135798854-7bd8a427-66a5-4e49-8d83-b85ba29d4c9c.png"> 

## Results
Thanks to the analysis we found out that in 2017 DQ grew almost 200% with a daily outcome of $35M. But in 2018 instead of increasing the stock price it shrunked by 63%. Compared to the other companies it seems it could had been a generalized situation; because only ENPH and RUN had a positive preformance of with more than 80%.

## Summary
Refactoring the script has been such a challenge, even if it has been able to reuse parts of previous work to save some time; at the end it got confused. Therefore it has been helpful having different versions of the script or rewrittings parts in a separate document.

Beeing not able to complete the callenge as expected makes us theoretically infer whether there is any disadvantage on the refactored script; but as stated before it was solution to make the VBA script run faster. Therefore it would have been an improvement to the original version.
