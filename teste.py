import locale
locale.getlocale()
('en_US', 'UTF-8')

locale.setlocale(locale.LC_TIME, 'pt-BR') # this sets the date time formats to es_ES, there are many other options for currency, numbers etc. 

import datetime
today = datetime.datetime.now()
today

datetime.datetime(2020, 2, 14, 10, 33, 56, 487228)

print(today.strftime('%A %d de %B, %Y'))

'viernes 14 de febrero, 2020'