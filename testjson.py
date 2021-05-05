import json
import plotly.express as px
from plotly.express import data

filename = 'file/eq_data_1_day_m1.json'
with open(filename) as f:
    all_data = json.load(f)

all_dicts = all_data['features']
print(len(all_dicts))

mags,titles,lons,lats = [],[],[],[]
for eq_dict in all_dicts:
    mag = eq_dict['properties']["mag"]
    title = eq_dict['properties']['title']
    lon = eq_dict['geometry']['coordinates'][0]
    lat = eq_dict['geometry']['coordinates'][1]
    titles.append(title)
    lons.append(lon)
    lats.append(lat)
    mags.append(mag)

print(mags[:10])
print(titles[:10])
print(lats[:10])
print(lons[:10])

fig = px.scatter(
    # data,
    x=lons,
    y=lats,
    labels={'x':'longitude','y':'latitude'},
    range_x= [-200,200],
    range_y=[-90,90],
    width=800,
    height=800,
    title="Seismogram",
    # size='level',
    # size_max=10,
)

fig.write_html('Seismogram.html')
fig.show()

# goal_json = 'file/goal_json.json'
# with open(goal_json,'w') as gf:
#     json.dump(all_data, gf, indent=4)

