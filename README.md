# -Coffe Sales Dashboard Excel Project


we are analyzing coffee order data to find trends in sales over time and by different coffee bean types, sales, by country, and we can gain some insights into who our top customers are.
The dashboard itself will be interactive, so you will be able to filter by time period, type of coffee roast, coffee sizes, and whether or not the customers are loyalty members.
My aim in this project is to explain, step-by-step my methodology and thought process.  So this particular project will not just focus on the end product, but the steps it took to get there.
 
Data Analysis Process
To start off, let’s have a look at our raw data:
Orders Table
On our Orders worksheet, we have columns A-E populated, but columns F-M are blank.  In order to populate this worksheet with the necessary data, we will be using lookups later on.  Those will be documented below.
Customers Table
Our customers table includes information you might expect on each of your customers, including their name, email, phone number, address, and loyalty member status.  Each row is identified by a unique customer ID (primary key).
Products Table
 
Similar to the customer table, the Product has a Product ID that functions as its primary key.  This table also contains useful information of the coffee type, roast type, size, unit price, price per 100g, and profit. 
 
Step 1: Gathering Data
First we must get all of the information from the latter two tables onto our Orders table.  In order to do that, we’ll utilize two different functions: XLOOKUP and INDEX/MATCH.
 
XLOOKUP
In order to populate the Customer Name column.
=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$B$1:$B$1001,,0)
The lookup value for this row is going to be the Customer ID (column C) in the Orders table.  I am going to refer to the cell in the same row for each individual value added.  As for a reference, I am using the Customer ID column (A) on the Customers table.  Since we want to return the customer’s name that is matched to the unique customer ID, the return array is the Customer Name column (B) on the Customers table.  We also want an exact match, so the [match mode] is set to ‘0’.
I repeated this process for the Email and Country columns in our Orders table.
=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)
And I ended up with values of ‘0’ in cells that did not return an email address.
I can remove the ‘0’ value by adding an IF to our original formula.
=IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0) = 0, "",XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0))
Basically, if the XLOOKUP returns a value of zero, I want the space to be blank, and if there is an email address, return that email address.  Turns out, this worked and the ‘0’ values were removed:
INDEX/MATCH
For the rest of the table, we will populate it using an INDEX/Match formula:
=INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0))
This formula can populate the Coffee Type, Roast type, Size, and Unit Price columns on the Orders table simply by filling.

Additional Formatting
The next step is to populate the Sales column (M) with a  simple formula, multiplying the Quantity by the Unit Price:
Ex:  =$L2*$E2
I also added full name columns for the coffee type and roast type:
Coffee Type Formula: =IF(I2="Rob", "Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica", IF(I2= "Lib", "Liberica",""))))
Roast Type Formula: =IF(J2="L", "Light", IF(J2= "M", "Medium", IF(J2= "D", "Dark", "")))
 
I then added some additional formatting for the Size, Unit Price, and Sales columns.
This process was repeated for the Country column (H):
=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$G$1:$G$1001,,0)
Finally, I converted the data to a table so it will be easier to work with when creating Pivot Tables.  This will make it much easier to create the final visualizations when the time comes.
 
Step 2: Pivot Tables
Build Total Sales Over Time
I created a pivot table by selecting the following parameters to get the total sales for each year and month between 2019 and 2022 using the following paraments in the Pivot Tables Fields Section:
Which results in this chart:
 
 
For the sake of consistency,  I decided to choose a theme here, and I developed my first pivot chart, which looks at the Total Sales Over Time.  This will later be added to the dashboard.
 
I also added a timeline filter for end users to operate.
Slicers
To add further customization for the end users, I also included slicers for
·       Roast Type Name
·       Size
·       Loyalty Member Status
These are useful to identify which coffee roasts are most popular, which sizes bring in the most, and help stakeholders decide whether it is worth it to promote their Loyalty Member program.
Sales By Country
Next is to create the bar chart for sales by country.  I duplicated the pivot table sheet and adjust the parameters of the pivot table to include the Country and sum of the sales:
Using this data, I created a bar graph for our final dashboard.  I added some formatting and ultimately ordered the sales in descending order.
Top 5 Customers
For our final chart, we want to identify the top five customers.  To do this. I once again duplicated my sheet to give me a baseline to start with.
To create a pivot table and chart, I simply swapped the Country with the Customer Name value, which led to a chart that is almost impossible to read:
The next step was to add the Top 5 filter to make this a much cleaner graphic.  I also wanted to sort my values from largest to smallest.
Now we have all of our elements completed, it’s time to build a dashboard that looks nice.
Step 3: Building a Dashboard
I started this by opening up a blank worksheet.  After copying all the information from the three pivot table sheets and some formatting, I finished with a  final dashboard that looks like this.
Finally, I had to adjust my slicers so that they updated for all three of my dashboard charts, not just the Total Sales Over Time.  This is done by going to ‘Timeline” and selecting ‘Report Connections’.  From there, I connected all three slicers to the charts.
Now, I can filter the charts using the timeline, roast type name, size, and loyalty member status.  With these filters, stakeholders can really hone in on what data they need given certain parameters.
After removing the gridline, voila!  The dashboard is complete!
 
 
Conclusion
There are certainly some insights you can make from the dashboard above.  
