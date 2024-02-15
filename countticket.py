
import pandas as pd
from utils import *

def countTicketReport():
  file_path = '/home/tspl/Documents/wsrAutomation/acp.xlsx'
  df_mirror = pd.read_excel(file_path, sheet_name='MIRRROR')

  end_date=current_sunday()
  start_date=current_monday()
  start_date = pd.to_datetime(start_date, format='%d/%m/%Y')
  end_date = pd.to_datetime(end_date, format='%d/%m/%Y')

  df_mirror['Date'] = pd.to_datetime(df_mirror['Date'])
  filtered_df = df_mirror[(df_mirror['Date'] >= start_date) & (df_mirror['Date'] <= end_date)]
  privity_counts = filtered_df['Priority'].value_counts().to_dict()

  result_dict = {'total': len(filtered_df), **privity_counts}
  print(result_dict)
  return result_dict


def countTicketReportMt():
  file_path = '/home/tspl/Documents/wsrAutomation/acp.xlsx'
  df_mirror = pd.read_excel(file_path, sheet_name='MODERN_TRADE')

  end_date=current_sunday()
  start_date=current_monday()
  start_date = pd.to_datetime(start_date, format='%d/%m/%Y')
  end_date = pd.to_datetime(end_date, format='%d/%m/%Y')

  df_mirror['Date'] = pd.to_datetime(df_mirror['Date'])
  filtered_df = df_mirror[(df_mirror['Date'] >= start_date) & (df_mirror['Date'] <= end_date)]
  privity_counts = filtered_df['Priority'].value_counts().to_dict()

  result_dict = {'total': len(filtered_df), **privity_counts}
  print(result_dict)
  return result_dict

