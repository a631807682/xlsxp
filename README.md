# xlsxp
Easier imports and exports with [xlsx](https://github.com/tealeg/xlsx)

## Description
> The `excel` tag defines what to export from `struct` to `XLXS`.

1. `header` represents the header to display.
2. `index` represents the order of data.
3. `format` represents a method of formatting.
4. `default` represents what to display when the data is empty.

## Example


### Export xlsx 
```go
type Test struct {
    UserName string `json:"user_name" excel:"header(Student);index(0);default(---)"`
	CompletePercent    float64 `json:"complete_percent" excel:"header(Complete Percent);index(1);format(percent)"`
}

datas := make([]Test, 0)
datas = append(datas, Test{
    UserName:        "A",
    CompletePercent: 15.3,
})

datas = append(datas, Test{
    UserName:        "B",
    CompletePercent: 17.558,
})

file, err := xlsxp.ExportExcel("sheet1", datas)
...

```

### Import xlsx 
```go
type Test struct {
    CompletePercent float64 `json:"complete_percent" excel:"index(1);parse(percent)"`
    UserName        string  `json:"user_name" excel:"index(0);"`
}

targetDatas := make([]Test, 0)
err = xlsxp.ImportExcel(xlsxBuf.Bytes(), "sheet1", &targetDatas)
```