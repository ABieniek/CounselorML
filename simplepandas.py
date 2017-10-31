import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlsmod

data = pd.read_csv('GradeDataNew/AllGrades.txt', sep="\t", header=None, low_memory = False)

'''data = pd.read_csv('GradeDataNew/AllGrades.txt', sep="\t", header=None,
    dtype={"Term": str, "Subject": str, "CRN": int, "Course Title": str,
        "Total Grades": int, "A+": int, "A": int, "A-": int, "B+": int, "B": int, "B-": int,
        "C+": int, "C": int, "C-": int, "D+": int, "D": int, "D-": int, "F": int, "W": int,
        "Average Grade": float, "Primary Instructor": str})'''

# xlsmod.headers

#print data
#print data.dtypes
#print data.describe()
#print data[2]

CS225Data = data.copy()
CS225Data = CS225Data[CS225Data[2].isin(['225.0'])]
CS225Data = CS225Data[CS225Data[1].isin(['CS'])]
print CS225Data
