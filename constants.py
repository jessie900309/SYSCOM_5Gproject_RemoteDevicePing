
COLUMN = ['C', 'D', 'E', 'F', 'G', 'H',
          'I', 'J', 'K', 'L', 'M', 'N',
          'O', 'P', 'Q', 'R',
          'S', 'T', 'U', 'V',
          'W', 'X', 'Y', 'Z',
          'AA', 'AB', 'AC', 'AD',
          'AE', 'AF', 'AG']

OuO_x23 = "SELECT DATE_FORMAT(`ReplyDateTime`, '%Y-%m-%d') AS rt, PingIP, `Status`, COUNT(*) FROM {tb} WHERE ReplyDateTime >= @st AND ReplyDateTime < @et AND DATE_FORMAT(`ReplyDateTime`, '%H') <> 23 GROUP BY rt, PingIP, `Status` ORDER BY rt, PingIP, `Status`;"
OuO_o23 = "SELECT DATE_FORMAT(`ReplyDateTime`, '%Y-%m-%d') AS rt, PingIP, `Status`, COUNT(*) FROM {tb} WHERE ReplyDateTime >= @st AND ReplyDateTime < @et GROUP BY rt, PingIP, `Status` ORDER BY rt, PingIP, `Status`;"

