Welcome to my command-line based schedule generator. This is specifically aimed at university schedules (and even more specifically at University of Guelph schedules, so other schedules may not work properly).

The program uses the following method to parse courses:

```0: course, 1: code, 2: section, 3: id, 4: lec, 5: start, 6: end, 7: freq, 8...N: day, N+1: lab, N+2: start, N+3: end, N+4: freq, N+5...N2: day, N+5+N2

s = 	engg, 2410, 01011, 4617, lec, 1300, 1420, 2, Tuesday, Thursday, lab, 0830, 1020, 1, Friday, sem, 1430, 1520, 1, 	Wednesday
	0     1     2      3     4    5     6     7  8        9         10   11    12    13 14      15   16    17    18 	19
	i     i     i      i     i    i     i     i  i+1      i+1       
									i+s[1]+1	 i+s[1]+4 	 i+s[i+s[1]+4]+2	i+[i+s[i+s[1]+4]]
									     i+s[1]+2	    i+s[i+s[1]+4]      i+s[i+s[1]+4]+3
										   i+s[1]+3	    i+s[i+s[1]+4]+1  i+s[i+s[1]+4]+4

s[0] = engg
s[1] = 2410
s[2] = 01011
s[3] = 4617
s[4] = lec
s[5] = 1300
s[6] = 1420
s[7] = 2
for i in range(1,s[7] = 2):
	i = 1	s[7+i] = s[7+1] = s[8] = Tuesday
	i = 2	s[7+i] = s[7+2] = s[9] = Thursday
s[8+s[7]] = s[8+2] = s[10] = lab
s[9+s[7]] = s[9+2] = s[11] = 0830
s[10+s[7]] = s[10+2] = s[12] = 1020
s[11+s[7]] = s[11+2] = s[13] = 1
for i in range(1,s[11+s[7]] = 1):
	i = 1	s[11+s[7]+i] = s[11+2+1] = s[14] = Friday

s[12+s[7]+s[11+s[7]]] = s[12+2+s[11+2]] = s[14+s[13]] = s[14+1] = s[15] = sem
s[13+s[7]+s[11+s[7]]] = s[13+2+s[11+2]] = s[15+s[13]] = s[15+1] = s[16] = 1430
s[14+s[7]+s[11+s[7]]] = s[14+2+s[11+2]] = s[16+s[13]] = s[16+1] = s[17] = 1520
s[15+s[7]+s[11+s[7]]] = s[15+2+s[11+2]] = s[17+s[13]] = s[17+1] = s[18] = 1
for i in range(1,s[15+s[7]+s[11+s[7]]] = 1):
	i = 1	s[15+s[7]+s[11+s[7]]+i] = s[15+2+s[11+2]+1] = s[17+s[13]+1] = s[17+1+1] = s[19] = Wednesday
```

Here's what it looks like in operation:

![Initial Run View Example](/screenshots/screenshot-1.png)
![List of Schedules Example](/screenshots/screenshot-2.png)
![Selection View Example](/screenshots/screenshot-3.png)

It automatically places the given schedules into a formatted excel spreadsheet, similar to the one below:

![Final Resultant Spreadsheet Example](/screenshots/screenshot-4.png)