package xlsxp

import (
	"fmt"
	"reflect"
	"time"
)

// 格式化方法
type FormatFn = func(fieldVal interface{}) string

type FormatFnMap = map[string]FormatFn

// 公共的格式化函数
var commonFormatFnMap = FormatFnMap{
	"percent": formatPercent,
	"date":    formDate,
	"hour":    formHour,
}

type CustomFormat struct {
	Name   string
	Format FormatFn
}

// 自定义格式化方法
func NewFormatFn(cformats ...CustomFormat) FormatFnMap {
	res := commonFormatFnMap
	if len(cformats) > 0 {
		for _, f := range cformats {
			res[f.Name] = f.Format
		}
	}
	return res
}

// 格式化百分比
func formatPercent(n interface{}) string {
	switch n.(type) {
	case int, int8, int16, int32, int64:
		if polyfillIsZero(reflect.ValueOf(n)) {
			return "0"
		}
		return fmt.Sprintf("%d%%", n)
	case float32, float64:
		if polyfillIsZero(reflect.ValueOf(n)) {
			return "0"
		}
		return fmt.Sprintf("%.2f%%", n)
	default:
		return ""
	}
}

// 格式化日期
func formDate(n interface{}) string {
	switch n.(type) {
	case time.Time:
		return n.(time.Time).Format("2006-01-02")
	case string:
		return n.(string)
	default:
		return ""
	}
}

func formHour(n interface{}) string {
	var seconds int64 = 0
	switch n.(type) {
	case int:
		seconds = int64(n.(int))
		break
	case int8:
		seconds = int64(n.(int8))
		break
	case int16:
		seconds = int64(n.(int16))
		break
	case int32:
		seconds = int64(n.(int32))
		break
	case int64:
		seconds = n.(int64)
		break
	}

	s := seconds % 60
	m := seconds / 60 % 60
	h := seconds / 3600
	return fmt.Sprintf("%d时%d分%d秒", h, m, s)
}
