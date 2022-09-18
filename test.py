import pandas as pd
from datetime import datetime


a = '20220101'
a = datetime.strptime(a, '%Y%m%d').strftime('%Y-%m-%d')
print(a)
print()