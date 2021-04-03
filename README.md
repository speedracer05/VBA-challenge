# VBA-challenge

This is an Excel VBA script for the VBA challenge assignment that analyzes the provided historical stock market data, summarizing yearly change between the opening and close price for a give year, percent change and total  volume for each unique stock. Formatting is also applied; conditional format to the yearly change, highlighting a cell green for positive, and red for negative changes; numerical formatting for percent change, and thousands formatting for the total stock volume.

The Bonus section was attempted, but not completed as time ran out. I'll attempt to complete on my own time --just for the challenge.

## Getting Started

The script is saved in a .bas file, which can be imported into the Excel VBA challenge data project.

### Prerequisites

You will need to use Excel, as well as the Excel file, "Multiple_year_stock__data.xlsx file, and the VBA file, "VBA_Stock_Challenge_jac.bas. 

### Installing

Using Excel, open the file "Multiple_year_stock__data.xlsx file
Click on the developer tab
Click on the tab "Visual Basic", and locate the File Explorer panel on left-panel of the VBA widow. 
Right-click the VBA Project "Multiple_year_stock_data" and select "import file" from the drop down menu
Select and open the file "VBA_Stock_Challenge_jac.bas. 

## Running the script

To run the script, you can select Run in the VBA Application window

The script will run through each worksheet tab and create a summary table in the worksheet; columns I to L. An example of the summary data is provided below.

Ticker  |  Yearly Change  |  Percent Change  | Total Stock Volume
 A              3.75                8.97%         528,574,200


## Built with
* Excel 2019 MSO (16.13801.20288) 64-bit
* [Maven](https://maven.apache.org/) - Dependency Management
* [ROME](https://rometools.github.io/rome/) - Used to generate RSS Feeds

## Contributing


## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **John Chan**

## Acknowledgments

* Hat tip to Bryan Tang for pointing me in the right direction on how to figure out the yearly change.
* Used information from Free Soft Dev as inspiration in the development of the VBA. https://freesoft.dev/program/163047389
