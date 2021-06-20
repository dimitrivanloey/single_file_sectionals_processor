import pandas as pd
import numpy as np

VENUE = 'hamilton'

DAY = '16'
MONTH = '06'

HOURS_LIST = ['1405', '1515']
TITLES_LIST = ['Title 1', 'Title 2']

FILE = f'HAM_{DAY}{MONTH}_sheets.xlsx'
DATE = '6th February 2021'

for n in range(len(HOURS_LIST)):
    df = pd.read_excel(FILE, sheet_name=HOURS_LIST[n])
    rounded_df = df.round({'Last 3f (%)': 2})
    new_df = rounded_df.replace(np.nan, '', regex=True)
    html_table = new_df.to_html(index=False, classes="horses", border=0)

    TIME = HOURS_LIST[n]

    html_string = '''
    <!DOCTYPE html>
<html>
<head>
<title>Post Race Reports</title>
<style>

</style>
<link type="text/css" rel="stylesheet" href="style.css">
</head>
<body>
   <div class="all">
       <div class="header1">
           <div class="race-logo"><img src="logos/{venue}.png" alt="Kempton Logo" style="width:180px;height:80px;">
           </div>
           <div class="title-date-container">
               <h2>{title}</h2>
               <h3>{date}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Time: {time}</h3>
           </div>
           <div class="racingTV-logo"><img src="racingTV.png" alt="Racing TV logo"
               style="width:300px;height:81px;">
           </div>
       </div>

    {table}

   </div>
   <div class="lower-banner">
       <div class="times-accuracy">
           <p>Times accurate to +/- 0.2s or more excluding the location accuracy +/-0.5m (approximately 0.03s).
               </br>One horse length is run in approximately 0.167 seconds on good or firmer ground, 0.18 seconds on
               good to soft ground and 0.2 seconds or more on soft.</p>
       </div>
       <div class="email" style="text-align:right;"><a href="mailto:sectionals@racingtv.com">For enquiries, please contact sectionals@racingtv.com</a>
       </div>
   </div>
</div>
</body>
</html>

    '''

    with open(f'{VENUE}_{HOURS_LIST[n]}.html', 'w') as website_file:
        website_file.write(
            html_string.format(table=new_df.to_html(index=False, classes="horses", border=0), title=TITLES_LIST[n],
                               date=DATE, time=TIME, venue=VENUE))














