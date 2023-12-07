import secrets
import string
import random
import pandas as pd

guests = []

for i in range(20):
	guests.append(''.join(secrets.choice(string.ascii_uppercase + string.ascii_lowercase) for i in range(7)))


tables = []
for i in range(20):
	tables.append(random.randint(1,30))

d = {'guest': guests, 'table': tables}
df = pd.DataFrame(data=d)

df.to_csv('guest_list.csv', index=False)