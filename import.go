package xlsxp

import (
	"fmt"
	"reflect"
	"strconv"

	"github.com/tealeg/xlsx"
)

// 导入excel
func ImportExcel(data []byte, sheetName string, targets interface{}, cparses ...CustomParse) (err error) {
	file, err := xlsx.OpenBinary(data)
	if err != nil {
		return
	}

	sheet, ok := file.Sheet[sheetName]
	if !ok {
		err = fmt.Errorf("sheet name:%s not exist", sheetName)
		return
	}

	// 初始化格式化函数
	parseFnMap := NewParseFn(cparses...)

	// 写入结构体
	err = setExcelKeyVals(sheet, targets, parseFnMap)
	if err != nil {
		return
	}

	return
}

type fieldInfo struct {
	fieldIndex int
	cellIndex  int
	parseFn    ParseFn
}

// 映射进数组
func setExcelKeyVals(sheet *xlsx.Sheet, targets interface{}, parseFnMap ParseFnMap) (err error) {
	targetsValue := reflect.ValueOf(targets)
	targetsInd := reflect.Indirect(targetsValue)

	if targetsValue.Kind() != reflect.Ptr || targetsInd.Kind() != reflect.Slice {
		err = fmt.Errorf("datas not slice ptr")
		return
	}

	// 列序号对应的字段序号
	cellFieldInfoMap := make(map[int]fieldInfo)
	itemType := targetsInd.Type().Elem()
	for fieldIndex := 0; fieldIndex < itemType.NumField(); fieldIndex++ {
		excelTag := itemType.Field(fieldIndex).Tag.Get(defaultStructTagName)
		if excelTag != "" { // 只处理定义了 `excel:"xxx"`
			_, tags := parseStructTag(excelTag)
			if headerIndex, ok := tags[tagIndex]; ok {
				cellIndex, _ := strconv.Atoi(headerIndex)

				fInfo := fieldInfo{
					fieldIndex: fieldIndex,
					cellIndex:  cellIndex,
				}

				if fnName, ok := tags[tagParse]; ok { //格式化值
					fn, ok := parseFnMap[fnName]
					if !ok {
						err = fmt.Errorf("excel tag parse func not exist, check common parse func or use defined custom func")
						return
					}

					fInfo.parseFn = fn
				}

				cellFieldInfoMap[cellIndex] = fInfo
			}
		}
	}

	for rIndex, row := range sheet.Rows {
		if rIndex == 0 {
			continue
		}

		// 创建一个item
		var elem reflect.Value
		typ := targetsInd.Type().Elem()
		if typ.Kind() == reflect.Ptr {
			elem = reflect.New(typ.Elem())
		}
		if typ.Kind() == reflect.Struct {
			elem = reflect.New(typ).Elem()
		}

		// 把列写入数据中
		for cIndex, cell := range row.Cells {
			fieldInfo, ok := cellFieldInfoMap[cIndex]
			if ok {
				targetField := elem.Field(fieldInfo.fieldIndex)
				if fieldInfo.parseFn != nil {
					pVal := fieldInfo.parseFn(cell.String())
					cell.SetValue(pVal)
				}

				setStrByKind(targetField, cell)
			}
		}

		targetsInd.Set((reflect.Append(targetsInd, elem)))
	}

	return
}

func setStrByKind(filed reflect.Value, cell *xlsx.Cell) {
	switch filed.Kind() {
	case reflect.Bool:
		filed.SetBool(cell.Bool())
		break
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		val, _ := cell.Int64()
		filed.SetInt(val)
		break
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		val, _ := strconv.ParseUint(cell.String(), 10, 64)
		filed.SetUint(val)
		break
	case reflect.Float64, reflect.Float32:
		val, _ := cell.Float()
		filed.SetFloat(val)
		break
	case reflect.String:
		filed.SetString(cell.String())
		break
	}
}
