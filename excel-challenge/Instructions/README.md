# Excel Homework: Charting Crowdfunding

## Background

Crowdfunding platforms like Kickstarter and Indiegogo have been growing in success and popularity since they began in the late aughts. Everyone from indie creators to famous celebrities have utilized crowdfunding to launch new products and generate buzz, but not every project has found success.

Getting funded on a crowdfunding website requires meeting or exceeding the project's initial goal, so many organizations spend months looking through past projects in an attempt to discover some trick for finding success. For this week's homework, you will organize and analyze a database of 1,000 generated sample projects in order to uncover any hidden trends.

### Before You Begin

1. Create a new space for this project called `excel-challenge` in either DropBox or Google Drive. **Do not add this homework to an existing space**.

2. Store your Excel workbooks here in this new space, and create a sharable link for submission.

## Instructions

![Crowdfunding Table](Images/FullTable.PNG)

Using the Excel [workbook](CrowdfundingBook.xlsx) provided, modify and analyze the data of 1,000 example projects in an attempt to uncover market trends. 

* Dataset created by Trilogy Education Services, LLC.


* Use conditional formatting to fill each cell in the `outcome` column with a different color, depending on whether the associated campaign was successful, failed, canceled, or is currently live.

  * Create a new column called `Percent Funded` that uses a formula to uncover how much money a campaign made relative to its initial goal.

* Use conditional formatting to fill each cell in the `Percent Funded` column according to a three-color scale. The scale should start at 0 and be a dark shade of red, transitioning to green at 100, and blue at 200.

  * Create a new column called `Average Donation` that uses a formula to uncover how much each project backer paid on average.

  * Create two new columns, one called `Parent Category` and another called `Sub-Category`, that use formulas to split the `Category and Sub-Category` column into the two new, separate columns.

  ![Category Stats](Images/CategoryStats.PNG)

  * Create a new sheet with a pivot table that will analyze your initial worksheet to count how many campaigns were successful, failed, canceled, or are currently live per **category**.

  * Create a stacked column pivot chart that can be filtered by country based on the table you have created.

  ![Subcategory Stats](Images/SubcategoryStats.PNG)

  * Create a new sheet with a pivot table that will analyze your initial sheet to count how many campaigns were successful, failed, or canceled, or are currently live per **sub-category**.

  * Create a stacked column pivot chart that can be filtered by country and parent-category based on the table you have created.

* The dates stored within the `deadline` and `launched_at` columns use Unix timestamps. Fortunately for us, [there is a formula](https://www.extendoffice.com/documents/excel/2473-excel-timestamp-to-date.html) that can be used to convert these timestamps to a normal date.

  * Create a new column named `Date Created Conversion` that will use [this formula](https://www.extendoffice.com/documents/excel/2473-excel-timestamp-to-date.html) to convert the data contained within `launched_at` into Excel's date format.

  * Create a new column named `Date Ended Conversion` that will use [this formula](https://www.extendoffice.com/documents/excel/2473-excel-timestamp-to-date.html) to convert the data contained within `deadline` into Excel's date format.

  ![Outcomes Based on Launch Date](Images/LaunchDateOutcomes.PNG)

  * Create a new sheet with a pivot table with a column of `outcome`, rows of `Date Created Conversion`, values based on the count of `outcome`, and filters based on `parent category` and `Years`.

  * Now create a pivot chart line graph that visualizes this new table.

* Create a report in Microsoft Word and answer the following questions.

1. Given the provided data, what are three conclusions we can draw about crowdfunding campaigns?
2. What are some limitations of this dataset?
3. What are some other possible tables and/or graphs that we could create, and what additional value would they provide?

## Bonus

* Create a new sheet with 8 columns:

  * `Goal`
  * `Number Successful`
  * `Number Failed`
  * `Number Canceled`
  * `Total Projects`
  * `Percentage Successful`
  * `Percentage Failed`
  * `Percentage Canceled`

* In the `Goal` column, create 12 rows with the following headers:

  * Less than 1000
  * 1000 to 4999
  * 5000 to 9999
  * 10000 to 14999
  * 15000 to 19999
  * 20000 to 24999
  * 25000 to 29999
  * 30000 to 34999
  * 35000 to 39999
  * 40000 to 44999
  * 45000 to 49999
  * Greater than or equal to 50000

  ![Goal Outcomes](Images/GoalOutcomes.PNG)

* Using the `COUNTIFS()` formula, count how many successful, failed, and canceled projects were created with goals within the ranges listed above. Populate the `Number Successful`, `Number Failed`, and `Number Canceled` columns with this data.

* Add up each of the values in the `Number Successful`, `Number Failed`, and `Number Canceled` columns to populate the `Total Projects` column. Then, using a mathematical formula, find the percentage of projects that were successful, failed, or canceled per goal range.

* Create a line chart that graphs the relationship between a goal amount and its chances of success, failure, or cancellation.

## Bonus Statistical Analysis

Most people would use the number of campaign backers to assess the success of a crowdfunding campaign. Creating a summary statistics table is one of the most efficient ways that data scientists can characterize quantitative metrics, such as the number of campaign backers.

For those of you looking for an additional challenge, evaluate the number of backers of successful and unsuccessful campaigns by creating **your own** summary statistics table.

* Create a new worksheet in your workbook, and create one column for the number of backers of successful campaigns and one column for unsuccessful campaigns.

  ![Images/backers01.png](Images/backers01.png)

* Use Excel to evaluate the following for successful campaigns, and then do the same for unsuccessful campaigns:

  * The mean number of backers

  * The median number of backers

  * The minimum number of backers

  * The maximum number of backers

  * The variance of the number of backers

  * The standard deviation of the number of backers

* Use your data to determine whether the mean or the median summarizes the data more meaningfully.

* Use your data to determine if there is more variability with successful or unsuccessful campaigns. Does this make sense? Why or why not?

## Submission

* To submit your homework, upload the solution and files to a GitHub repo, Dropbox, or Google Drive, and submit the link to <https://bootcampspot.com/>.

## Rubric

[Unit 1 Rubric - Excel Homework: Charting Crowdfunding](https://docs.google.com/document/d/1gIVky_5CZi-b4w07RLBYZyXgbbhMRemvpcT1V6EGkvw/edit?usp=sharing)

## Employer-Ready Criteria

Students who are marked as employer ready gain access to our employer referral program, additional workshops, and other resources. Work with your Career Director to become employer ready. At a minimum, you must have:

- A clear, concise, and compelling resume. Submit via your learning platform for review.
- A polished GitHub profile:
  - 3 to 6 pinned repositories ([instructions here](https://docs.github.com/en/enterprise/2.13/user/articles/pinning-items-to-your-profile))
  - Professional titles, i.e. not "Homework #1"
  - Thorough README.md files for each repository
  - Clean code

## References

Dataset created by Trilogy Education Services, LLC.

- - -

© 2021 Trilogy Education Services, a 2U, Inc. brand. All Rights Reserved.