# jsonxlsx

## Executables

Compile the package with the flag `exe` to get two executables, namely
`json2xlsx` and `xlsx2json`.
Their usage is illustrated below.
See `json2xlsx --help` and `xlsx2json --help` to get additional information.

- Write a file with `json2xlsx`:

```
> json2xlsx -d {\"A\":[1,2],\"B\":[3,\"x\"]} --header -o outfile.xlsx
> json2xlsx -d {\"A\":[1],\"B\":[3,\"x\"]} --header -o outfile.xlsx
> json2xlsx -d {\"A\":[null,2],\"B\":[3,\"x\"]} --header -o outfile.xlsx
```

- Read it with `xlsx2json`:

```
> xlsx2json -f outfile.xlsx -s Sheet1 -w data --header
{"A":[null,2],"B":[3,"x"]}
```

Here the flag `data` tells to `xlsx2json` to read the cell values.

You can read the cell types:

```
> xlsx2json -f outfile.xlsx -s Sheet1 -w types --header
{"A":[null,"number"],"B":["number","text"]}
```

Or both the cell values and the cell types:

```
> xlsx2json -f outfile.xlsx -s Sheet1 -w "data,types" --header
{"data":{"A":[null,2],"B":[3,"x"]},"types":{"A":[null,"number"],"B":["number","text"]}}
```

You can specify a first row and a last row:

```
> xlsx2json -f outfile.xlsx -s Sheet1 -w data -F 2 -L 10
{"X1":[null,2],"X2":[3,"x"]}
```

- Write and read some comments:

```
> json2xlsx -d {\"A\":[null,2],\"B\":[3,\"x\"]} -c "{\"A\":[\"a comment\",null],\"B\":[null,\"another one\"]}" --header -o outfile.xlsx
```

```
> xlsx2json -f outfile.xlsx -s Sheet1 -w comments --header
{"A":["a comment",null],"B":[null,"another one"]}
```

It is also possible to read the cell formats:

```
> xlsx2json -f outfile.xlsx -s Sheet1 -w formats --header
{"A":[null,null],"B":[null,null]}
```

- Read all worksheets

Ignore the flag `-s` to read all sheets of a `xlsx` file:

```
> xlsx2json -f myfile.xlsx -w data --header
```
