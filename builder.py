import pandas as pd
import xlwt

excel_file = '191213_Schedule.xlsx'

# General parameters
parameters_gen = pd.read_excel(excel_file, sheet_name='Parameters')
parameters_gen.head()

days_count = parameters_gen['Value'][0]
shift_count = parameters_gen['Value'][1]
shift_length = parameters_gen['Value'][2]
cell_size = parameters_gen['Value'][3]
shift_cells = parameters_gen['Value'][4]

# Cycle parameters
parameters_cycle = pd.read_excel(excel_file, sheet_name='Cycle')
parameters_cycle.head()

# Ops list initialization
ops = []

ops.append(parameters_cycle["# of process"].to_list())
ops.append(parameters_cycle["Length, cells"].to_list())
ops.append(parameters_cycle["Parallel site work"].to_list())
ops.append(parameters_cycle["Partial completion"].to_list())

print(ops[0])
print(ops[1])


class ScheduleNode:
    def __init__(self, shift_size):
        self.first_site = []
        self.second_site = []

        for i in range(shift_size):
            self.first_site.append(99)
            self.second_site.append(99)

        self.first_site_exp = 0
        self.second_site_exp = 0
        self.first_site_cont = 0
        self.second_site_cont = 0

    def print_node(self):
        print('1st site: ', self.first_site, '1st site explosion: ', self.first_site_exp)
        print('2nd site: ', self.second_site, '2nd site explosion: ', self.second_site_exp)


# Data initialization
schedule = [[ScheduleNode(shift_cells) for shift in range(shift_count)] for day in range(days_count)]
# print(len(schedule), len(schedule[0]))

# Cycle completions parameters
first_site_seq_len = 0
first_site_next_sh = 0

prev_sh_day = 0
prev_sh_sh = 0

for day in range(days_count):
    for shift in range(shift_count):
        for i in range(shift_cells):
            # Set first operation
            if i == 0:

                # Shift after explosion
                if ((day == 0) and (shift == 0)) \
                    or ((day >= 0) and (shift == 1) and (schedule[day][shift - 1].first_site_exp == 1)) \
                        or ((day > 0) and (shift == 0) and (schedule[day-1][shift+1].first_site_exp == 1)):
                    schedule[day][shift].first_site[i] = ops[0][0]
                    first_site_seq_len = 1

                # Other days
                else:
                    if shift == 1:
                        prev_sh_day = day
                        prev_sh_sh = shift - 1
                    else:
                        prev_sh_day = day - 1
                        prev_sh_sh = shift + 1
                    if schedule[prev_sh_day][prev_sh_sh].first_site_exp == 1:
                        schedule[day][shift].first_site[i] = ops[0][0]

                    elif schedule[prev_sh_day][prev_sh_sh].first_site_cont == 0:
                        # Find last action on previous day except 99
                        j = shift_cells-1

                        if schedule[prev_sh_day][prev_sh_sh].first_site[j] == 99:
                            while schedule[prev_sh_day][prev_sh_sh].first_site[j] == 99:
                                j -= 1
                        schedule[day][shift].first_site[i] = \
                            ops[0][ops[0].index(schedule[prev_sh_day][prev_sh_sh].first_site[j]+1)]
                        first_site_seq_len = 1

                    elif schedule[prev_sh_day][prev_sh_sh].first_site_cont == 1:
                        schedule[day][shift].first_site[i] = \
                            schedule[prev_sh_day][prev_sh_sh].first_site[shift_cells-1]
                        first_site_seq_len += 1

            # Set other operations
            else:
                if (ops[1][ops[0].index(schedule[day][shift].first_site[i - 1])] - first_site_seq_len) > 0:
                    if (ops[1][ops[0].index(schedule[day][shift].first_site[i - 1])] -
                            first_site_seq_len) <= (shift_cells - i):
                        schedule[day][shift].first_site[i] = schedule[day][shift].first_site[i - 1]
                        first_site_seq_len += 1

                elif ops[0][ops[0].index(schedule[day][shift].first_site[i - 1])] < ops[0][len(ops[0]) - 1]:

                    if ops[0][ops[0].index(schedule[day][shift].first_site[i - 1])] == ops[0][len(ops[0]) - 2]:
                        schedule[day][shift].first_site_exp = 1
                        schedule[day][shift].first_site[i] = 99

                    elif ops[1][ops[0].index(schedule[day][shift].first_site[i - 1] + 1)] < (shift_cells - i):
                        schedule[day][shift].first_site[i] = \
                            ops[0][ops[0].index(schedule[day][shift].first_site[i - 1] + 1)]
                        first_site_seq_len = 1

                    elif ops[3][ops[0].index(schedule[day][shift].first_site[i - 1] + 1)] == 1:
                        schedule[day][shift].first_site[i] = \
                            ops[0][ops[0].index(schedule[day][shift].first_site[i - 1] + 1)]
                        schedule[day][shift].first_site_cont = 1
                else:
                    schedule[day][shift].first_site[i] = 99
                    schedule[day][shift].first_site_cont = 0

# Print schedule
for day in range(days_count):
    for shift in range(shift_count):
        print('Day: ', day, 'Shift: ', shift)
        schedule[day][shift].print_node()

