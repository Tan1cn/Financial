
import imp
import pandas as pd
import numpy as np
from datetime import datetime


a = '20220101'
a = datetime.strptime(a, '%Y%m%d').strftime('%Y-%m-%d')
print(a)
