import pandas as pd

spring_sched = pd.read_csv('Spring_Division_Enrollment.csv')

spring_sched_subset = spring_sched.iloc[:, 3:18]

print(spring_sched_subset.to_string())
drops_df = pd.read_csv('Copy of Liberal Arts AHC and Undecided Spring 2023 ENR 01.03.2023.csv')
sorted_spring = spring_sched_subset.sort_values(by=['Class']).reset_index()
# print(sorted_spring.to_string())
# drops_df = drops_df[drops_df['Instruction Mode Description'] != 'Laboratory']
pd.set_option('display.max_columns', None)
drops_df_subset = drops_df[['Employee ID', 'Section Number', 'Course', 'Instruction Mode Description', 'Enrollment Add Date', 'Enrollment Drop Date']]
# sorted_drops = drops_df_subset.sort_values(by=['Section Number'])

for i in range(len(drops_df_subset)):
    for j in range(len(sorted_spring)):
        if drops_df_subset.loc[i, 'Section Number'] == sorted_spring.loc[j, 'Class']:
            drops_df_subset.loc[i, 'Instruction Mode Description'] = spring_sched.loc[j, 'Type']
        if 'Africana' in drops_df_subset.loc[i, 'Course']:
            drops_df_subset.loc[i, 'Instruction Mode Description'] == 'Lecture'


# print(drops_df_subset.to_string())




