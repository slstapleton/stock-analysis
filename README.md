# Yearly Analysis of Green Energy Stocks through VBA

## Overview of Project

### Purpose
Steve has been collecting stock market data on various green energy companies. To track the success and determine wise investments, Steve wants to find the total daily volume, number of shares traded in a day, and the yearly return, percent difference in price from the beginning to the end of a given year. Steve would like to use this VBA program for possible future analysis of stocks and wants to find the most efficient submacro to run large volumes of data as fast as possible. 

## Analysis

### Analysis of 2017 Green Energy Stocks Information
The year 2017 seemed very successful for green energy. All but one stock, TERP, had positive returns on investment which is easily determined with any positive percentage appearing green in color on our table. Looking at total daily volume we can determine how active an account is. Some believe the higher the volume, the more accurately the return on investment is represented. In this case, FSLR had the second highest total daily volume and one of the highest returns, so this stock might be worth investing in, followed by ENPH and SEDG.

### Analysis of 2018 Green Energy Stocks Information
2018 green energy stocks dropped and resulted in negative returns for most stocks which is easily determined with any negative percentage appearing red in color on our table. There were two stocks with positive returns and a drastic increase in total daily volume, ENPH and RUN. Based on the data provided for the 2017 and 2018, ENPH may be the best option for investing; however, looking into other economic factors prior to investing would be a wise choice.

## Results

- What are the advantages or disadvantages of refactoring code?
There are many advantages to refactoring code; one of which being the main focus of this challenge, efficiency. With multiple small macros there may be repetition and clutter. By creating a single submacro, this increases the readability and overall neatness of the code. Refactoring also lessens the chances of bugs and is formatted for ease of future use and development. In regards to the green stocks analysis, we wanted to create a single button to run a series of commands into a single output. Buttons can only be assigned one macro as a task, so in order to have one button to complete the entire table, the small set of submacros had to be combined.

In our original coding, we have the formatting portion laid out in it’s own submacro:
![original_2](https://user-images.githubusercontent.com/100329223/159199076-d4176aa2-20db-46dd-8813-6253327adfce.png)

Below is an example of how we could simply remove the end and start of of a new submacro, as well as multiple sheet activations, through our refactored code:
![refactored_2](https://user-images.githubusercontent.com/100329223/159199101-807d6512-417b-45cf-baf1-e2644509b908.png)

In our original coding is listed below showing variables, the number of rows to loop through, and a for loop:
![original_1](https://user-images.githubusercontent.com/100329223/159199131-3056a989-1c06-42b0-bc55-80a285ecda1c.png)

To alter this section, we removed the clutter and reassigned our arrays and added a ticker index:
![refactored_1](https://user-images.githubusercontent.com/100329223/159199138-a9bf80d3-854c-4bf5-9040-cb99fbb5634c.png)

A disadvantage could be that each submacro is not clearly defined and laid out step by step. Having multiple small commands, you could go straight to the title of the submacro to determine what that section of coding accomplishes instead of searching through an all-inclusive submacro. Refactoring could also be very time consuming if you aren’t familiar with the sequence of coding. Altering the code order could create errors in the output that are difficult to locate and trouble shoot.


- How do these pros and cons apply to refactoring the original VBA script?
In this particular case with the analysis of green stocks, we were able to compute the time it took to run the submacro prior to refactoring and after.  Prior to refactoring, the run times were 0.96875 seconds for 2017, and 0.94528 seconds for 2018. Some changes that were made included reducing the number of submacros and removing any lines of repetition/clutter. Following the changes, we were able to reduce the run time of these tables to under 0.9 seconds as seen below. In the future, Steve can use this file of coding to analyze larger files with ease and minimal changes.

