package xlsxp

import "strings"

// 解析方法
type ParseFn = func(cellVal string) string

type ParseFnMap = map[string]ParseFn

// 公共的解析函数
var commonParseFnMap = ParseFnMap{
	"percent": parsePercent,
}

// 格式化百分比
func parsePercent(val string) string {
	return strings.TrimSuffix(val, "%")
}

type CustomParse struct {
	Name  string
	Parse ParseFn
}

// 自定义解析方法
func NewParseFn(cparses ...CustomParse) ParseFnMap {
	res := commonParseFnMap
	if len(cparses) > 0 {
		for _, f := range cparses {
			res[f.Name] = f.Parse
		}
	}
	return res
}
