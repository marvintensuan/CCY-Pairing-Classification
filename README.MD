# Disclaimer

First of all, I would like to state that the data in these files are purely a product of fiction and are generated for educational purpose only.

# Rationale of the Problem

Let me introduce *currency pairs*.

Every trade in the FX Market involves two currencies. For example you gave up US$1 (USD SELL) for a CA$1.3 (CAD BUY). That is a currency pair.

Say, in a work setting where there are segragation of duties, there could be a lot of people working across different teams that have different interests for different tasks. For example, a person inputting cash inflows and outlows in the database may not care about what kind of cash these are. But you do.

Hence, whenever you pull data from the database, you want to determine which of those are currency pairs. You have been matching them manually but now, you decided to make a script to automate the classification.


# Data

The Excel file contains a script that showcase a classification algorithm.

On the upper region of the worksheet (`B1:E5`) are the data from the foreign currencies you bought for August. These are classified as "First", "Second", "Third", and "Fourth".

There exists a USD Equivalent column as an accounting requirement. Long story short, we can see that it is different, but it should give us a feel on how close it could be to how much it was traded. For simplicity's sake, we can call it the *"booking rate"*.

On the lower region (`B9:E16`) are data from US$ you have sold. You need to fill the `Classification` column to determine which of them would pair with "First", "Second", "Third", and "Fourth", respectively.

Notice that there are more than one USD entry on the 7th of August, but only one should match with "Second". Also, take note that there are two $10M SELL entries, but the "Fourth" pair should match on the entry recorded on 25-Aug.

The `Classify` button will run the script you developed.

(For those of you who are wary of downloading Excel files on the internet, the script is also available as a BAS file)

# Solution

Although it could be debatable whether or not the script generated falls under the machine learning domain, it is certainly inspired by it.

The *"cost"* of each *"feature"* were pitted against each output.

The `cost` is expressed as the sum of the cost of features `f1` and `f2`.

The first cost of feature `f1` is the squared difference of the number of days between the recorded dates. The second cost of feature `f2` is the squared difference between the booking rate and the actual amount, scaled to the booking rate (or more simply, square of the percentage change from the booking rate).

Each cost is stored in `memo` array and the index of the minimum value will provide the cell to give the output.

(No actual memoization were made in the making of the script).

What this solution doesn't have are the so-called *weights* or *biases*. It would have came in handy when say on 7-Aug, there is another entry of US$ 5,620,000. Or when you have an entry on 8-Aug of US$ 5,624,500. Right now, let's assume we have a bias `theta` = [0 1 1]'.

Unfortunately, performing matrix operations on Excel VBA is something that's currently not within my skillset. Perhaps one day, it would be feasible to develop a gradient descent algorithm on VBA.

