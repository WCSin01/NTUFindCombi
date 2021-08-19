import sys
import pandas as pd
import itertools

#Assume no same course at same timing but different week
schedule_df = pd.read_excel("Schedule.xlsx", usecols="A:B, E:F, H")

#Modify data for analysis
#Index is renamed CI (class index) to avoid ambiguity
schedule_df.rename(columns={"Index": "CI"}, inplace=True)
#ci_row lists starting "index" of each CI
ci_rows = schedule_df["CI"].dropna().index.tolist()
#course_rows lists starting "index" of each course
course_rows = schedule_df["Course"].dropna().index.tolist()
course_names = schedule_df["Course"].dropna().tolist()

day_to_date =\
    {"Mon": "1970-01-05", "Tue": "1970-01-06", "Wed": "1970-01-07", "Thu": "1970-01-08", "Fri": "1970-01-09"}
schedule_df["Day"].replace(day_to_date, inplace=True)
temp = schedule_df["Time"].str.split("to", expand=True)
schedule_df["Start Time"], schedule_df["End Time"] = temp[0], temp[1]
schedule_df["Start Time"], schedule_df["End Time"] =\
    schedule_df["Day"] + " " + schedule_df["Start Time"],\
    schedule_df["Day"] + " " + schedule_df["End Time"]
schedule_df["Start Time"], schedule_df["End Time"] =\
    pd.to_datetime(schedule_df["Start Time"]), pd.to_datetime(schedule_df["End Time"])

schedule_df["CI"] = schedule_df["CI"].ffill().astype(int)
schedule_df.drop(["Course", "Day", "Time"], axis=1, inplace=True)


start_df = schedule_df.sort_values(by="Start Time").reset_index()
end_df = schedule_df.sort_values(by="End Time").reset_index()
schedule_df["Clashes"] = ""

#Find clashes
for course_row in course_rows[1:]:
    course_df = start_df.drop(start_df[start_df["index"] >= course_row].index).reset_index(drop=True)
    start_df = start_df.drop(start_df[start_df["index"] < course_row].index).reset_index(drop=True)
    end_df = end_df.drop(end_df[end_df["index"] < course_row].index).reset_index(drop=True)

    start_insert_indexes = start_df["Start Time"].searchsorted(course_df["Start Time"])
    end_insert_indexes = end_df["End Time"].searchsorted(course_df["End Time"])
    i = 0
    for start_insert_index, end_insert_index in zip(start_insert_indexes, end_insert_indexes):
        #"index" between start_insert_index (inclusive) and start_clash_index (exclusive) clashes.
        #selected end time is later than start times.
        #"index" between end_clash_index (inclusive) and end_insert_index (exclusive) clashes.
        #selected start time is earlier than end times.
        #column 3 is start time, column 4 is end time.
        start_clash_index =\
            start_df.iloc[start_insert_index:, 3].searchsorted(course_df.at[i, "End Time"]) + start_insert_index
        end_clash_index =\
            end_df.iloc[:end_insert_index, 4].searchsorted(course_df.at[i, "Start Time"])
        schedule_df.at[course_df.at[i, "index"], "Clashes"] =\
            start_df.iloc[start_insert_index:start_clash_index, 0].tolist() +\
            end_df.iloc[end_clash_index:end_insert_index, 0].tolist()
        i = i + 1

#Combine clashes into first timing of each CI
first = 0
for last in ci_rows[1:]:
    if last - first > 1:
        #column 4 is clashes.
        schedule_df.iat[first, 4] = schedule_df.iloc[first:last, 4].sum()
        schedule_df.iloc[first+1:last, 4] = ""
    first = last

#Find viable combinations
ci_row_by_course = []
first = 0
for last in course_rows[1:]:
    ci_row_by_course.append(ci_rows[first:ci_rows.index(last)])
    first = last
ci_row_by_course.append(ci_rows[ci_rows.index(last):])

viable_combis_row = []
for combi in list(itertools.product(*ci_row_by_course)):
    try:
        clashes = set(schedule_df.iloc[list(combi[:-1]), 4].sum())
        if set(combi).intersection(clashes) == set():
            viable_combis_row.append(list(combi))
    except TypeError:
        viable_combis_row.append(list(combi))

#Check if there are no viable combinations
if viable_combis_row == []:
    holder = input("No viable combinations. Hit Enter to continue...")
    sys.exit()

#Rank combinations
combi_df = pd.Series(viable_combis_row, name="Combi").to_frame()

def scorer(combi):
#column 2 and 3 are start and end time
    temp = schedule_df.iloc[combi, 2:4].sort_values(by="Start Time")
#find duration between classes
    temp = temp.iloc[1:, 0].reset_index(drop=True) - temp.iloc[:-1, 1].reset_index(drop=True)
#total duration between classes on same day
    duration = temp[temp < pd.Timedelta(11, unit="h")].sum() / pd.Timedelta(1, unit="min")
#day count
    day = temp[temp > pd.Timedelta(11, unit="h")].shape[0] + 1
    return duration, day

combi_df["Duration"], combi_df["Day"] = zip(*combi_df["Combi"].map(scorer))
combi_df["Duration"] = combi_df["Duration"].round(3)
trpt = int(input("One-way transport time (in min): "))
combi_df["Duration and Trpt"] = combi_df["Duration"] + trpt*2*combi_df["Day"]

#Export viable combinations
#Expand combi into separate columns for each course
combi_expand_df = pd.DataFrame(combi_df["Combi"].tolist(), columns=course_names)
#Convert index to CI
i_to_ci_dict = schedule_df["CI"].to_dict()
combi_expand_df.replace(i_to_ci_dict, inplace=True)

combi_df = pd.concat([combi_df["Duration and Trpt"], combi_df["Day"], combi_df["Duration"], combi_expand_df], axis=1)
combi_df.sort_values(["Duration and Trpt", "Day"], inplace=True)
combi_df.rename(columns={"Duration": "Mins betw classes on same day per week",
                         "Day": "No. of days of week with classes",
                         "Duration and Trpt": "Mins betw classes on same day + trpt time"}, inplace=True)

#xlsxwriter
writer = pd.ExcelWriter("Viable Combinations.xlsx", engine="xlsxwriter")
#Write df
combi_df.to_excel(writer, startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets["Sheet1"]
#Write wrap header
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
})
for col_num, value in enumerate(combi_df.columns.values):
    worksheet.write(0, col_num, value, header_format)
#set column width
worksheet.set_column(0, 2, 10)
writer.save()

holder = input("Process finished. Hit Enter to continue...")