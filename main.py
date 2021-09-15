import pandas as pd
import numpy as np
import webbrowser
import os

VENUE = 'sandown'

DAY = '10'
MONTH = '09'

HOURS_LIST = ['1300', '1330', '1400', '1430', '1500', '1535']
TITLES_LIST = ["0m 5f 10y Paul Ferguson Memorial EBF Maiden Stakes (GBB Race)",
               "0m 5f 10y 1account ID Handicap",
               "1m 0f 0y IRE Incentive Scheme EBF Fillies' Novice Stakes (GBB Race)",
               "1m 0f 0y 1account KYC Handicap",
               "0m 7f 0y Follow Raceday On Instagram Handicap",
               "0m 7f 0y Every Race Live On Racing TV Fillies' Handicap"
               ]

FILE = f'SAN_{DAY}{MONTH}.xlsx'
DATE = '10th September 2021'

for n in range(len(HOURS_LIST)):
    df = pd.read_excel(FILE, sheet_name=HOURS_LIST[n])
    rounded_df = df.round({'Last 3f (%)': 2})
    new_df = rounded_df.replace(np.nan, '', regex=True)
    html_table = new_df.to_html(index=False, classes="horses", border=0)

    TIME = HOURS_LIST[n]
    HH = TIME[0:2]
    MM = TIME[2:4]

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
           <div class="race-logo"><img src="logos/{venue}.png" alt="{venue} Logo" style="width:180px;height:80px;">
           </div>
           <div class="title-date-container">
               <h2>{title}</h2>
               <h3>{date}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Time: {hh}:{mm}</h3>
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
                               date=DATE, time=TIME, venue=VENUE, hh=HH, mm=MM))

    filename = 'file:///' + os.getcwd() + '/' + f'{VENUE}_{HOURS_LIST[n]}.html'
    webbrowser.open_new_tab(filename)












