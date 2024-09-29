
'''
def rec_fact(n):
    if n == 0:
        return 0
    else:
        return n%10 + rec_fact(n//10)
    
print(rec_fact(6721))
'''

import pandas as pd

data = {
"Name": ["Emma","Gireeja", "Sophia"],
"Age": [15, 28, 22],
"City": ["Dubai","London", "San Jose"]
}

df = pd.DataFrame(data)

df
print(df)

