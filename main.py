import pandas as pd
from support import (add_flag,get_col,generate_custom_excel,generate_excel,create_df,append_all_data,mark_duplicate_in,remove_duplicate_pair,formatting_multiple_shift,
                    mark_duplicate_out,get_colored_formatted_excel,append_cleaned,append_duplicate,append_all_data,manually_verify_pair,main_code)

if __name__ == '__main__':
    filename =   input("  Enter the Filename : ")
    start_date = input("  Enter the startdate format is month/date/year  11/1/2022 : " )
    end_date   = input("  Enter the end date format is month/date/year 11/30/2022  : ")
    main_code(filename)

df_lst , buffer_hour_time ,shift_hour_int , man_verify_hour_int ,buffer_hour_int= main_code(filename)
marked_out = mark_duplicate_out(df_lst , buffer_hour_time)
marked_in_out = mark_duplicate_in(df_lst , buffer_hour_time)
duplicate_data = append_duplicate(marked_in_out)
no_duplicate_data = append_cleaned(marked_in_out)
proper_pair = append_all_data(no_duplicate_data)

manually_verify = manually_verify_pair(proper_pair,shift_hour_int , man_verify_hour_int,buffer_hour_int)
no_duplicate_df = pd.DataFrame(no_duplicate_data,columns=["LogDateTime","punch_code","emp_code","emp_name","Company","Department","date","time","flag"])
proper_pair_df = pd.DataFrame(proper_pair,columns=["emp_code","emp_name","Company","Department","punch_code","in_date",
                                                        "in_time","punch_code","out_date","out_time","muster_time","hrs"])
proper_pair_df = add_flag(proper_pair_df)
x = proper_pair_df["hrs"].fillna(0)
proper_pair_df["hrs"] = x
proper_pair_df_lst = proper_pair_df.values.tolist()
proper_pair_new = formatting_multiple_shift(proper_pair_df_lst)
proper_pair , overtime_pair= remove_duplicate_pair(proper_pair_new)
overtime_pair_df , proper_pair_df , manually_verify_df , duplicate_df = create_df(overtime_pair,proper_pair, manually_verify , duplicate_data)
generate_excel(no_duplicate_df , manually_verify_df , manually_verify_df , duplicate_df)
unique_emp_code = proper_pair_df["emp_code"].unique()
col = ['Emp Code','Emp Name','Company','Designation']
dates = pd.date_range(start=start_date,end=end_date)
column = get_col(col , dates)
df1 = generate_custom_excel(unique_emp_code , column , overtime_pair_df , manually_verify_df,proper_pair_df,start_date , end_date)
get_colored_formatted_excel(df1,overtime_pair_df,manually_verify_df)
'''

in and out both missing then we use red color
in but no out or out but no in then yellow
overtime - blue color

'''