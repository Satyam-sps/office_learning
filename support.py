import pandas as pd
import numpy as np
from icecream import ic
from datetime import datetime
def seperate_date_time(data):
    """
    This function returs to list which contains all the date and time
    """
    date = []
    time = []
    for x in range(len(data)):
        data_ , time_ = data['Log Date'][x].split(" ")
        date.append(data_)
        time.append(time_)
    return date , time

def convert_str_time(time):
    """
    takes in input a string in the format of Hr:Min:Sec and convert it to time
    """
    time = datetime.strptime(time,"%H:%M:%S")
    return time

def str_time(time):
    """
    takes as input string and format it to day:month:year Hr:Min:Sec
    """
    time = datetime.strptime(time,"%d-%b-%Y %H:%M:%S")
    return time



def convert_timedelta_to_hour(time):
    """
    Args:
        time (sec) : This is an integer type argument
    Return:
        It return Hours:minutes:sec and hours 
    """

    hours, remaider = divmod(time , 3600)
    minutes , seconds  = divmod(remaider , 60)
    return f"{int(hours)}:{int(minutes)}:{int(seconds)}" , int(hours)

def main_code(filename):
    df = pd.read_excel(filename)
    df = df.loc[7:]
    df.drop(columns=['Unnamed: 0','Unnamed: 4','Unnamed: 6','Unnamed: 8','Unnamed: 10'],inplace=True)
    df.rename(columns={'Log Records Report (Device Wise)':'Log Date','Unnamed: 2':'punch_code','Unnamed: 3':'emp_code',
                        'Unnamed: 5':'emp_name','Unnamed: 7':'Company','Unnamed: 9':'Department'},inplace=True)
    df = df.reset_index().drop(columns='index')
    df = add_date_time_col(df)
    df_lst  = df.values.tolist()
    for x in range(len(df_lst)):
        format_time_date = str_time(df_lst[x][0])
        df_lst[x][0] = format_time_date
    buffer_hour = '7:00:00'
    buffer_hour_int = convert_str_time(buffer_hour)
    shift_hour = '8:00:00'
    shift_hour_int = convert_str_time(shift_hour)

    test_hour = '0:00:00'
    test_hour_int = convert_str_time(test_hour)
    man_verify_hour = '15:00:00'
    man_verify_hour_int = convert_str_time(man_verify_hour)
    
    shift_hour_time = (shift_hour_int - test_hour_int)
    buffer_hour_time = buffer_hour_int - test_hour_int
    return df_lst , buffer_hour_time,shift_hour_int , man_verify_hour_int,buffer_hour_int

def mark_duplicate_out(dummy_data,shift_hour_time):
    """
    This function  start looping from bottom to top and it sees if we find the next entry to be out and date of both entries is same and differences
    between the two datetime is less than shift_hour_time then we mark the next entry as 1 ie it is duplicate
    """
    for x in range(len(dummy_data)-1 , -1 ,-1):
        current_entry = dummy_data[x]
        if current_entry[1] == 'out':
            for y in range(x-1 ,-1 , -1 ):
                next_entry = dummy_data[y]
                time_delta = current_entry[0] - next_entry[0]
                if next_entry[1] == 'out' and current_entry[2] == next_entry[2] and current_entry[6] == next_entry[6] and time_delta< shift_hour_time: 
                    dummy_data[y][8] = '1'
    return dummy_data

def mark_duplicate_in(dummy_data ,shift_hour_time):
    """
    IF the next entry is in , date of both entries is same and datetime differnces between two entries if it is less than shift hour time then 
    we mark the next entry as 1 ie it is a duplicate in 
    """
    for x in range(len(dummy_data)):
        current_entry = dummy_data[x]
        if current_entry[1] == 'in':
            for y in range(x+1,len(dummy_data)):
                next_entry = dummy_data[y]
                time_delta = next_entry[0] - current_entry[0]
                if next_entry[1] == 'in' and current_entry[2] == next_entry[2] and current_entry[6] == next_entry[6] and time_delta < shift_hour_time:
                    dummy_data[y][8] = '1'
    return dummy_data

def append_duplicate(data):
    """
    Here data consists of the data which flag column contains 0 and 1 1 means the data is duplicated 0 means the data is not duplicated so we check if row is duplicated
    if so we the append the row to duplicate
    """
    duplicate_ = []
# To store the duplicate in array
    for x in range(len(data)):
        if data[x][8] == '1':
            duplicate_.append(data[x])
    return duplicate_


def append_cleaned(dummy_data_after_in_out):
    """
    Here data consists of the data which flag column contains 0 and 1 0 means the data is not duplicated and 1 means data is duplicated , so now 
    we check if row is duplicated if so we the append the row to duplicate
    """
    cleaned_data = []
    for x in range(len(dummy_data_after_in_out)):
        if dummy_data_after_in_out[x][8] == '0':
            cleaned_data.append(dummy_data_after_in_out[x]) 
    return cleaned_data


def append_all_data(cleaned_data):
    """
    cleaned_data 

    """
    all_data = []
    for x in range(len(cleaned_data)-1):
        current_entry = cleaned_data[x]
        current_data = current_entry[1]
        next_entry = cleaned_data[x+1]
        next_data = next_entry[1]
        #print(current_entry)
        if (current_data == "in" and next_data == "out") and (current_entry[2] == next_entry[2]):
            muster_time = next_entry[0]-current_entry[0]
            sec = muster_time.total_seconds()
            hour_min_sec , hours= convert_timedelta_to_hour(sec)
            all_data.append([current_entry[2],current_entry[3],current_entry[4],current_entry[5],current_entry[1],current_entry[6],current_entry[7],next_entry[1],next_entry[6],next_entry[7],hour_min_sec , hours])
        if (current_data == 'in' and next_data == 'in') and (current_entry[2] == next_entry[2]):
            all_data.append([current_entry[2],current_entry[3],current_entry[4],current_entry[5],current_entry[1],current_entry[6],current_entry[7],np.nan,np.nan , np.nan, np.nan, 0])
    # Made changes from -1 to 0 and the out pair went for KUMAR MALI
    for x in range(len(cleaned_data)-1,0,-1):
        current_entry = cleaned_data[x]
        next_entry = cleaned_data[x-1]
        current_date = current_entry[1]
        next_date = next_entry[1]
        if (current_date == 'out' and next_date == 'out') and (current_entry[6] != next_entry[6]):
            all_data.append([current_entry[2],current_entry[3],current_entry[4],current_entry[5],np.nan,np.nan,np.nan,next_entry[1],current_entry[6],current_entry[7],np.nan,0])
    return all_data

def manually_verify_pair(proper_pair,shift_hour_int , man_verify_hour_int,buffer_hour_time):
    manually_verify = []
    for x in range(len(proper_pair)):
        shift_hour = shift_hour_int.hour
        man_verify = man_verify_hour_int.hour
        buffer_hr = buffer_hour_time.hour
        hour = proper_pair[x][11]
        # if (hour < shift_hour) or (hour > shift_hour + shift_hour/2):
        #     manually_verify.append(proper_pair[x])
        #print(hour , shift_hour)
        if (hour < buffer_hr) or (hour > man_verify):
            manually_verify.append(proper_pair[x])
        if pd.isna(hour):
            manually_verify.append(proper_pair[x])
    return manually_verify

def add_flag(proper_pair_df):
    for x in range(len(proper_pair_df)):
        proper_pair_df["flag"] = '0'
    return proper_pair_df


def create_df(overtime_pair , proper_pair, manually_verify , duplicate_data):
    overtime_pair_df = pd.DataFrame(overtime_pair,columns=["emp_code","emp_name","Company","Department","punch_code","in_date",
                                                            "in_time","punch_code","out_date","out_time","muster_time","hrs","flag"])
    overtime_pair_df.drop(columns='flag',inplace=True)

    proper_pair_df = pd.DataFrame(proper_pair,columns=["emp_code","emp_name","Company","Department","punch_code","in_date",
                                                            "in_time","punch_code","out_date","out_time","muster_time","hrs","flag"])

    manually_verify_df = pd.DataFrame(manually_verify,columns=["emp_code","emp_name","Company","Department","punch_code","in_date",
                                                            "in_time","punch_code","out_date","out_time","muster_time","hrs"])
    duplicate_df = pd.DataFrame(duplicate_data,columns=["LogDateTime","punch_code","emp_code","emp_name","Company","Department","date","time","flag"])
    return overtime_pair_df , proper_pair_df , manually_verify_df , duplicate_df



def generate_excel(no_duplicate_df , proper_pair_df , manually_verify_df , duplicate_df):
    no_duplicate_df.to_excel('no_duplicate_data.xlsx')
    proper_pair_df.to_excel('proper_pair.xlsx')
    manually_verify_df.to_excel('manually_verify.xlsx')
    duplicate_df.to_excel('duplicate_data.xlsx')

def in_out_hrs(emp_instance):
    data = []
    for x in range(len(emp_instance)):
        if (pd.isna(emp_instance[x][5]) and  pd.notna(emp_instance[x][6])):
            hrs = 'NA'
            data.append(emp_instance[x][5])
            data.append(emp_instance[x][6])
            data.append(hrs)
        elif (pd.isna(emp_instance[x][6]) and pd.notna(emp_instance[x][5])):
            hrs = 'NA'
            data.append(emp_instance[x][5])
            data.append(emp_instance[x][6])
            data.append(hrs)
        else:    
            data.append(emp_instance[x][5])
            data.append(emp_instance[x][6])
            data.append(emp_instance[x][7])
    emp_header = [emp_instance[0][0],emp_instance[0][1],emp_instance[0][2],emp_instance[0][3]]
    emp_header.extend(data)
    return emp_header

def all_dates_data(start_date , end_date , emp_instance_lst):
    dates = pd.date_range(start=start_date,end=end_date)
    dates_format = []
    for x in range(len(dates)):
        y = dates[x].strftime('%d-%b-%Y')
        dates_format.append(y)
    list_ = []
    for x in range(len(dates_format)):
        date_x = dates_format[x]
        flag = 0
        for y in range(len(emp_instance_lst)):
            record = emp_instance_lst[y]
            in_date_y = emp_instance_lst[y][5]
            if date_x == in_date_y:
                flag = 1
                list_.append([record[0],record[1],record[2],record[3],record[5],record[6],record[9],record[11]])
        if flag == 0:
            list_.append([record[0],record[1],record[2],record[3],date_x,np.nan,np.nan,0])
    return list_



def get_col(col , dates):
    format_dates = []
    for x in dates:
        a = x.strftime('%d-%b')+', In'
        format_dates.append(a)
        b = x.strftime('%d-%b')+', Out'
        format_dates.append(b)
        c = x.strftime('%d-%b')+', Hrs'
        format_dates.append(c)
    col.extend(format_dates)
    return col



def emp_row_data(unique_emp_code , proper_pair_df , start_date , end_date):   
    emp_row_data = [] 
    for x in unique_emp_code:
        emp_instance = proper_pair_df[proper_pair_df["emp_code"] == x]
        emp_instance_lst = emp_instance.values.tolist()
        all_data = all_dates_data(start_date ,end_date,emp_instance_lst)
        data = in_out_hrs(all_data)
        emp_row_data.append(data)
    return emp_row_data


def generate_custom_excel(unique_emp_code , column , overtime_pair_df , manually_verify_df,proper_pair_df,start_date , end_date):
    emp_row_data_ = emp_row_data(unique_emp_code,proper_pair_df,start_date , end_date)
    data = []
    for x in range(len(emp_row_data_)):
        data.append(emp_row_data_[x])
    df1 = pd.DataFrame(data,columns=column)
    a = df1.columns.str.split(', ',expand=True).values
    df1.columns = pd.MultiIndex.from_tuples([('', x[0]) if pd.isnull(x[1]) else x for x in a])
    return df1


def get_list(df1):
    return df1.values.tolist()

def append_cleaned_no_duplicate_in(cleaned_data):
    data = []
    for x in range(len(cleaned_data)-1):
        current_entry = cleaned_data[x]
        next_entry = cleaned_data[x+1]
        if current_entry[3] != next_entry[3]:
            data.append(current_entry) 
    return data



def add_date_time_col(df):
    date , time = seperate_date_time(df)
    df["date"] = date
    df["time"] = time
    flag_ = []
    for x in range(len(date)):
        flag_.append('0')
    df["flag"] = flag_
    return df


def formatting_multiple_shift(proper_pair):
    for x in range(len(proper_pair)-1):
        current_entry = proper_pair[x]
        next_entry = proper_pair[x+1]
        #print(current_entry[5] , next_entry[5])
        if current_entry[5] == next_entry[5]:
            next_entry[12] = '1'
    return proper_pair


def remove_duplicate_pair(proper_pair_new):
    all_data = []
    overtime_data = []
    for x in range(len(proper_pair_new)):
        current_entry = proper_pair_new[x]
        if current_entry[12] == '0':
            all_data.append(current_entry)
        if current_entry[12] == '1':
            overtime_data.append(current_entry)
    return all_data , overtime_data

def highlight_cells(val):
    if val == 'NA':
        color = 'yellow'
    elif val == 0:
        color = 'red'
    elif val > 12:
        color = 'lightblue'
    else:
        color = 'white'
    return f'background-color: {color}'
        
def get_colored_formatted_excel(df1,overtime_pair_df,manually_verify_df):
    x = df1.xs('Hrs',level=1,axis=1,drop_level=False)
    a = df1.style.applymap(highlight_cells,subset=x.columns)
    a.to_excel('test.xlsx')
    writer = pd.ExcelWriter('final.xlsx', engine='openpyxl')
    a.to_excel(writer, sheet_name='attendance_report')
    overtime_pair_df.to_excel(writer, sheet_name='Overtime Record')
    manually_verify_df.to_excel(writer, sheet_name='Manually Verify')
    writer.save()
    
    
