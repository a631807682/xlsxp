# xlsxp ![test](https://github.com/a631807682/xlsxp/actions/workflows/test.yml/badge.svg)
Easier imports and exports excel with [xlsx](https://github.com/tealeg/xlsx)

## Description
> The `excel` tag defines what to import and export from `struct` to `XLXS`.

### Import
1. `index` represents the order of data.
2. `parse` represents a method of parsing corresponds to the `format`.
### Export
1. `index` represents the order of data.
2. `header` represents the header to display.
3. `format` represents a method of formatting.
4. `default` represents what to display when the data is empty.
5. `width` represents the header width.

### Export xlsx example  
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

file := xlsx.NewFile()
err := xlsxp.ExportExcel(file, "sheet1", datas)
...
err := xlsxp.ExportExcel(file, "sheet2", datas)
...
file.Save(filepath)
...

```

### Import xlsx example  
```go
type Test struct {
    CompletePercent float64 `json:"complete_percent" excel:"index(1);parse(percent)"`
    UserName        string  `json:"user_name" excel:"index(0);"`
}

targetDatas := make([]Test, 0)
err = xlsxp.ImportExcel(xlsxBuf.Bytes(), "sheet1", &targetDatas)
```

### Notice
Not support huge amounts of data. See:  
[xlsx/issues/539](https://github.com/tealeg/xlsx/issues/539)  
[xlsx/blob/master/file.go#L169](https://github.com/tealeg/xlsx/blob/master/file.go#L169)  
[xlsx/blob/master/file.go#L410](https://github.com/tealeg/xlsx/blob/master/file.go#L410)  