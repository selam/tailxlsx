# tailxlsx

tailxlsx is go-lang application to provide see content of xlsx files (Microsoft Excel format) on console
## parameters of _tailxlsx_

`-s` sheet number or name default is "1" (first sheet)

`-r` start and end row numbers (no space between them), seperated with coma default is 1,10

`-c` start and end cell numbers, you must give cell number instead of letters, default is 1,10


# Usage of tailxlxs

tailxlsx test.xlsx
```
+----+----+----+
| 1a | 1b | 1c |
+----+----+----+
| 2a | 2b | 3c |
+----+----+----+
```


### TODO

make it more readable

write tests

add odf support



# Thanks
for tablewriter to [Oleku Konko](http://github.com/olekukonko/tablewriter)

for xlsx library to [Geoffrey J. Teale](https://github.com/tealeg/xlsx)

