package xlsxp

import (
	"fmt"
	"strings"
)

const (
	defaultStructTagName  = "excel"
	defaultStructTagDelim = ";"
)

const (
	tagHeader  = "header"  //表头
	tagIndex   = "index"   //顺序
	tagDefault = "default" //默认值
	tagFormat  = "format"  //格式化
)

var supportTag = map[string]int{
	tagHeader:  2,
	tagIndex:   2,
	tagDefault: 2,
	tagFormat:  2,
}

func parseStructTag(data string) (attrs map[string]bool, tags map[string]string) {
	attrs = make(map[string]bool)
	tags = make(map[string]string)
	for _, v := range strings.Split(data, defaultStructTagDelim) {
		if v == "" {
			continue
		}
		v = strings.TrimSpace(v)
		if t := strings.ToLower(v); supportTag[t] == 1 {
			attrs[t] = true
		} else if i := strings.Index(v, "("); i > 0 && strings.Index(v, ")") == len(v)-1 {
			name := t[:i]
			if supportTag[name] == 2 {
				v = v[i+1 : len(v)-1]
				tags[name] = v
			}
		} else {
			fmt.Println("unsupport excel tag", v)
		}
	}
	return
}
