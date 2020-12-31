import datetime
date = 'june 1954'
dt = datetime.datetime.strptime( date, '%b %Y' )
if dt.year > 2000:
    dt = dt.replace( year=dt.year-100 )
print(dt.strftime( '%Y' ))