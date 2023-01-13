import pandas as pd
import openpyxl

fall_courses = pd.read_csv('Fall_Division_Enrollment.csv')
pd.set_option('display.max_columns', None)
fall_courses_subset = fall_courses[['Dept','Session', 'Class', 'Room']]
fall_courses_subset= fall_courses_subset.fillna(0)
fall_courses_subset = fall_courses_subset.astype({'Class': int})
print(fall_courses_subset.to_string())
fall_success = pd.read_csv('Fall 2022 Completion and Success.csv')
# fall_success = fall_success.astype({'Enrolled': int,'Success': int})
print(fall_success)
print(fall_success.dtypes)
# # print(fall_courses_subset.shape)
for i in range(len(fall_courses_subset)):
    for j in range(len(fall_success)):
        if fall_courses_subset.loc[i, 'Class'] == fall_success.loc[j,'Section']:
            fall_courses_subset.loc[i, 'Completion'] = fall_success.loc[j, 'Enrolled']
            fall_courses_subset.loc[i, 'Success'] = fall_success.loc[j, 'Success']
fall_courses_subset = fall_courses_subset.fillna(0)
fall_courses_subset = fall_courses_subset.astype({'Class': int, 'Completion': int, 'Success': int})
print(fall_courses_subset)
print(fall_courses_subset.dtypes)
ol_compl = 0
ol_succ = 0
ol_count = 0

na_compl = 0
na_succ = 0
na_count = 0

ip_compl = 0
ip_succ = 0
ip_count = 0

hyb_compl = 0
hyb_succ = 0
hyb_count = 0

rem_compl = 0
rem_succ = 0
rem_count = 0

for i in range(len(fall_courses_subset)):
    # print('Online', ol_compl, ol_succ, ol_count)
    if fall_courses_subset.loc[i, 'Room'] == 'In Person':
        ip_compl += fall_courses_subset.loc[i, 'Completion']
        ip_succ += fall_courses_subset.loc[i, 'Success']
        ip_count += 1
        # print('In Person', ip_compl, ip_succ, ip_count)
    elif fall_courses_subset.loc[i, 'Room'] == 'Hybrid':
        hyb_compl += fall_courses_subset.loc[i, 'Completion']
        hyb_succ += fall_courses_subset.loc[i, 'Success']
        hyb_count += 1
        # print('Hybrid', hyb_compl, hyb_succ, hyb_count)
    elif fall_courses_subset.loc[i, 'Room'] == 'REMOTE':
        rem_compl += fall_courses_subset.loc[i, 'Completion']
        rem_succ += fall_courses_subset.loc[i, 'Success']
        rem_count += 1
        # print('Remote', rem_compl, rem_succ, hyb_count)
    elif fall_courses_subset.loc[i, 'Room'] == 'ONLINE':
        ol_compl += fall_courses_subset.loc[i, 'Completion']
        ol_succ += fall_courses_subset.loc[i, 'Success']
        ol_count += 1
        # print('9A', hyb_compl, hyb_succ, hyb_count)

    else:
        na_compl += fall_courses_subset.loc[i, 'Completion']
        na_succ += fall_courses_subset.loc[i, 'Success']
        na_count += 1
        # print('Misc', ol_compl, ol_succ, ol_count)

eighteen_compl = 0
eighteen_succ = 0
eighteen_count = 0

fifteenA_compl = 0
fifteenA_succ = 0
fifteenA_count = 0

fifteenB_compl = 0
fifteenB_succ = 0
fifteenB_count = 0

nineA_compl = 0
nineA_succ = 0
nineA_count = 0

nineB_compl = 0
nineB_succ = 0
nineB_count = 0

six_compl = 0
six_succ = 0
six_count = 0

nan_compl = 0
nan_succ = 0
nan_count = 0

for i in range(len(fall_courses_subset)):

    if fall_courses_subset.loc[i, 'Session'] == '18':
        eighteen_compl += fall_courses_subset.loc[i, 'Completion']
        eighteen_succ += fall_courses_subset.loc[i, 'Success']
        eighteen_count += 1
        # print('18', eighteen_compl, eighteen_succ, eighteen_count)
    elif fall_courses_subset.loc[i, 'Session'] == '15A':
        fifteenA_compl += fall_courses_subset.loc[i, 'Completion']
        fifteenA_succ += fall_courses_subset.loc[i, 'Success']
        fifteenA_count += 1
        # print('fifteenA', fifteenA_compl, fifteenA_succ, fifteenA_count)
    elif fall_courses_subset.loc[i, 'Session'] == '15B':
        fifteenB_compl += fall_courses_subset.loc[i, 'Completion']
        fifteenB_succ += fall_courses_subset.loc[i, 'Success']
        fifteenB_count += 1
    elif fall_courses_subset.loc[i, 'Session'] == '9A':
        nineA_compl += fall_courses_subset.loc[i, 'Completion']
        nineA_succ += fall_courses_subset.loc[i, 'Success']
        nineA_count += 1
    elif fall_courses_subset.loc[i, 'Session'] == '9B':
        nineB_compl += fall_courses_subset.loc[i, 'Completion']
        nineB_succ += fall_courses_subset.loc[i, 'Success']
        nineB_count += 1
        # print('9B', nineB_compl, nineB_succ, nineB_count)
    elif fall_courses_subset.loc[i, 'Session'] == '6':
            six_compl += fall_courses_subset.loc[i, 'Completion']
            six_succ += fall_courses_subset.loc[i, 'Success']
            six_count += 1
            # print('6', six_compl, six_succ, six_count)
    else:
        nan_compl += fall_courses_subset.loc[i, 'Completion']
        nan_succ += fall_courses_subset.loc[i, 'Success']
        nan_count += 1
        # print('Misc', nan_compl, nan_succ, nan_count)

print(f'18, {eighteen_compl}, {eighteen_succ}, {eighteen_count} {eighteen_succ / eighteen_compl}')
print(f'15A, {fifteenA_compl}, {fifteenA_succ}, {fifteenA_count}, {fifteenA_succ / fifteenA_compl}')
print(f'15B, {fifteenB_compl}, {fifteenB_succ}, {fifteenB_count}, {fifteenB_succ / fifteenB_compl}')
print(f'9A, {nineA_compl}, {nineA_succ}, {nineA_count}, {nineA_succ / nineA_compl}')
print(f'9B, {nineB_compl}, {nineB_succ}, {nineB_count}, {nineB_succ / nineB_compl}')
print(f'6, {six_compl}, {six_succ}, {six_count}, {six_succ / six_compl}')
print('Nan', nan_compl, nan_succ, nan_count)

        # print('Remote', rem_compl, rem_succ, rem_count)
print(f'Online, {ol_compl}, {ol_succ}, {ol_count} {ol_succ / ol_compl}')
print(f'In Person, {ip_compl}, {ip_succ}, {ip_count}, {ip_succ / ip_compl}')
print(f'Hybrid, {hyb_compl}, {hyb_succ}, {hyb_count}, {hyb_succ / hyb_compl}')
print(f'Remote, {rem_compl}, {rem_succ}, {rem_count}, {rem_succ / rem_compl}')
print('Na', na_compl, na_succ, na_count)

fall_courses_subset.to_excel('Success_by_Modality.xlsx')