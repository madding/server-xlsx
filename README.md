# Simple server take params and return xslx file
- columns - column names
- data - data for xlsx
- data[number] - number it's random number for unique records

## Get request
```
  localhost:5000?columns[]=id&columns[]=name&data[1]=1&data[1]=yourname
```

## Post request
```
  curl localhost:5000 -X POST -d "columns[]=id" -d "columns[]=name" -d "data[1]=1" -d "data[1]=eu" -d "data[2]=2" -d "data[2]=ue"
```
