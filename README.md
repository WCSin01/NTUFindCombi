# NTUFindCombi
Generates schedule for NTU STARS and finds viable combinations ranked by minimal time wastage.

## Prerequisites
1. python
2. pip

## Instructions
1. Save your courses under Plan 3. It do not need to be a valid combination.
2. In cmd:
	```
	pip install -r requirements.txt
	python GenerateSchedule.py
	```
3. Key in your username and password for NTU LDAP. The script will not save your login information.
4. In cmd:
	```
	python FindCombi.py
	```
5. Key in your travel time.
6. Open Viable Combinations.xlsx

### Using the output
1. There will be MANY viable combinations.
2. They are sorted in ascending order by minutes between lessons on the same day and travel time per week ("rank").
	a. If you have a longer travel time, combinations with lessons on fewer days will be on top.
	b. If you have a shorter travel time, combinations with less time between lesssons will be on top, even if they are over more days.
3. You can also use the excel to sort the other 2 columns to your preferred order. E.g. if you stay in hall and want lessons on more days.
4. Among the many combinations with the same "rank", you can agree upon with your friends what all of you prefer so that you could share classes with them.
5. You can introduce this tool to your friends who share some courses but differ in others to find common index no. in both your top "ranked" combinations.

## Potential problems and fixes
1.
If the website structure changes, GenerateListing will not work. You can create your own schedule and continue from step 5.
The column headings are as follows: Course, Index, Type, Group, Day, Time, Venue, Remark
Only Course, Index, Day, Time, Remark are strictly necessary. The remaining columns can be blank columns.

2.
If the course does not have a schedule, GenerateSchedule will prompt and still create excel sheet, but will not include the course.

3.
If the course combination in the excel sheet does not have a viable combination, FindCombi will prompt.
Pls choose which courses you want to forgo and remove it, then run FindCombi again.
