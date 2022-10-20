# Stock-analysis
## Excel VBA Module 2

### OverView of Stock_Analysis (VBA_Challenge) Project

	The stock analysis project analyzes an entire dataset of stock information (volume traded, starting prices and ending prices of a certain, sometimes selected, stock in a given year).
The purpose behind the project was by using different VBA formulas and functions, we could analyze several key stocks and their related information, without the need to go over them one by one, which is essentially time consuming.
At a click of a button, the project would analyze one to tens of different stocks, at any given year, analyzing key data such as the return of an example stock.

### Results
	Stock perfomance in 2017 was better than 2018
Nearly all (10 / 12) stocks out performed themselves in 2017 over 2018. Only tickers (RUN) and (TERP) had lesser returns in 2017 than in 2018. 
This might indicate that, perhaps, the overall economy, or atleast the selected sector, had faced a decline and recession in 2018.
If we were to focus solely on "DQ" Ticker, as selected through module 2, we would see that it faced a huge decline on return from 2017 to 2018.
![This image is from 2017](https://github.com/AliBailoun234/Stock-analysis/blob/main/Extra_Resources/DQ_2017.png)
![This image is from 2018](https://github.com/AliBailoun234/Stock-analysis/blob/main/Extra_Resources/DQ_2018.png)
Also, the volume traded increased threefold from 2017 to 2018, perhaps from the decline "DQ" experienced.  


The same code was used to analyze both 2017 and 2018 years, in both the original and refactored scripts.

##### Original script vs refactored script
	Shorter run time, lesser code used and looks better.
The refactored script had a better execution time over the original script, as shown in these two images in 2017. 
![This image is from the refactored script](https://github.com/AliBailoun234/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![This image is from the original script](https://github.com/AliBailoun234/Stock-analysis/blob/main/Extra_Resources/2017_Original_Script.png)
The first image is from the refactored script, while the second image is from the original script. They clearly show how with a refactored script, execution time decreases significantly. 
This difference sometimes adds up to a difference of one second between the original and refactored script. 2018 also had a small execution time as shown in the following image.
![This image is from the refactored script](https://github.com/AliBailoun234/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
It took less than 0.2 seconds to analyze a list of stocks in 2018. 


### Summary
There are several advantages and disadvantages of refactoring code in general, some of which were discussed in the Module.
##### Advantages of refactoring
One advantage of refactoring is that it makes the code run faster, as shown previously in the analysis part. It also makes the code look better.
##### Disadvantages of refactoring
Disadvantages of refactoring might be is being risky. Because refactoring requires specific and attention to detail, one mistake, misspelling, might require hours of debugging.

#### Applying advantages and disadvantages to original and refactored code
As discussed previously, it makes the code run faster and makes the looks of the code better by making it shorter
But, for me atleast, refactoring the original code took way more time to write and work on over writing the original code, as there were immense details that I accidently overlooked while refactoring the original code.

