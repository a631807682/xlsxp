# xlsxp
Easier to use https://github.com/tealeg/xlsx

## Description
> 通过`excel`标签定义`struct`中要导出到`xlxs`的内容。

1. `header`表示要显示的表头。
2. `index`表示数据在`xlxs`中的顺序。
3. `format`表示格式化的方法，目前`format.go`文件中只定义了`percent`一种，可通过`cformats`参数定义私有的格式化方法，或在`format.go`中定义公有的方法。
4. `default`表示当[数据为空](https://golang.org/src/reflect/value.go?s=34297:34325#L1090)时要显示的内容。

## Example


### Export xlsx 
```go
type Test struct {
    UserName string `json:"user_name" excel:"header(学员);index(0);default(---)"`
	CompletePercent    float64 `json:"complete_percent" excel:"header(完课率);index(1);format(percent)"`
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