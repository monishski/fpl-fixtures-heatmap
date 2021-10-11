import pandas as pd
import xlwings as xw

rank = {
  'Man City': 1,
  'Man Utd': 1,
  'Liverpool': 1,
  'Chelsea': 1,
  'Leicester': 2,
  'West Ham': 2,
  'Spurs': 1,
  'Arsenal': 1,
  'Leeds': 2,
  'Everton': 2,
  'Aston Villa': 2,
  'Newcastle': 3,
  'Wolves': 2,
  'Crystal Palace': 3,
  'Southampton': 3,
  'Brighton': 3,
  'Burnley': 3,
  'Norwich': 3,
  'Watford': 3,
  'Brentford': 3
}

fixtures = {
  'Man City': [],
  'Man Utd': [],
  'Liverpool': [],
  'Chelsea': [],
  'Leicester': [],
  'West Ham': [],
  'Spurs': [],
  'Arsenal': [],
  'Leeds': [],
  'Everton': [],
  'Aston Villa': [],
  'Newcastle': [],
  'Wolves': [],
  'Crystal Palace': [],
  'Southampton': [],
  'Brighton': [],
  'Burnley': [],
  'Norwich': [],
  'Watford': [],
  'Brentford': []
}

color = {
  1: (255, 0, 0),
  2: (255, 165, 0),
  3: (0, 255, 0),
}

df = pd.read_csv('./data.csv')

for index, row in df.iterrows():
  fixtures[row['home']].append([row['away'] + ' (H)', rank[row['away']]])
  fixtures[row['away']].append([row['home'] + ' (A)', rank[row['home']]])

app = xw.App(visible=False)
wb = xw.Book()
wb.sheets['Sheet1'].name = 'data'
sheet = wb.sheets['data']

column = 'ABCDEFGHIJKLMNOPQRSTUVXYZ'

index = 0
for key, value in fixtures.items():
  sheet.range(column[index]+'1').value = key
  for i, [team, difficulty] in enumerate(value):
    sheet.range(column[index]+str(i+2)).value = team
    sheet.range(column[index]+str(i+2)).color = color[difficulty]
  index += 1

wb.save('./fixtures.xlsx')
wb.close()
app.quit()
