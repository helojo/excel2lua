package main

import (
	"bytes"
	"fmt"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

const (
	COMMENT = 0 // first line
	TYPE    = 1 // second line
	KEY     = 2 // third line
	// other is content
)

var SPLIT = []string{
	",", // first split sign
	";", // second split sign
	"|", // third split sign
}

func parse(input_filename string, output_filename string) {
	file, err := xlsx.OpenFile(input_filename)
	if err != nil {
		panic(err)
	}

	// Just parse first sheet
	sheet := file.Sheets[0]
	if sheet == nil {
		panic("can't find sheet")
	}

	var type_list []string
	var key_list []string

	buf := new(bytes.Buffer)
	buf.WriteString("return {\n")

	for row_id, row := range sheet.Rows {
		if row_id == COMMENT {
			// Just ignore
		} else if row_id == TYPE {
			for _, cell := range row.Cells {
				s, err := cell.String()
				if err != nil {
					panic(err)
				}
				type_list = append(type_list, s)
			}
		} else if row_id == KEY {
			for _, cell := range row.Cells {
				s, err := cell.String()
				if err != nil {
					panic(err)
				}
				key_list = append(key_list, s)
			}
		} else {
			// Content
			id, err := row.Cells[0].Int64() // first column reserved for id
			if err != nil {
				panic(err)
			}
			buf.WriteString(fmt.Sprintf("%s[%d] = {\n", padding(1), id))
			parse_col(buf, row, type_list, key_list, 2)
			buf.WriteString(fmt.Sprintf("%s},\n", padding(1)))
		}
	}

	buf.WriteString("}")

	out, err := os.Create(output_filename)
	if err != nil {
		panic(err)
	}

	out.Write(buf.Bytes())
	out.Close()
}

func parse_col(buf *bytes.Buffer, row *xlsx.Row, type_list []string, key_list []string, nest_level int) {
	for col_id := 1; col_id < len(row.Cells); col_id++ {
		parse_cell(buf, key_list[col_id], row.Cells[col_id], type_list[col_id], nest_level)
	}
}

func parse_cell(buf *bytes.Buffer, key string, val *xlsx.Cell, val_type string, nest_level int) {
	if key == "" || val_type == "" {
		return
	}

	val_type = strings.Trim(val_type, " ")
	vt_list := strings.Split(val_type, "_")

	if len(vt_list) <= 1 {
		parse_atom(buf, key, val, val_type, nest_level)
	} else if len(vt_list) <= 4 {
		val_type = vt_list[0]
		s, err := val.String()
		if err != nil {
			panic(err)
		}

		buf.WriteString(fmt.Sprintf("%s[%q] = {\n", padding(nest_level), key))
		parse_list(buf, key, s, val_type, nest_level, len(vt_list)-1)
		buf.WriteString(fmt.Sprintf("%s},\n", padding(nest_level)))
	} else {
		panic("only support three nest list")
	}
}

func parse_atom(buf *bytes.Buffer, key string, val *xlsx.Cell, val_type string, nest_level int) {
	if val_type == "string" {
		s, err := val.String()
		if err != nil {
			panic(err)
		}
		buf.WriteString(fmt.Sprintf("%s[%q] = %q,\n", padding(nest_level), key, s))
	} else if val_type == "integer" {
		i, err := val.Int64()
		if err != nil {
			i = 0
		}
		buf.WriteString(fmt.Sprintf("%s[%q] = %d,\n", padding(nest_level), key, i))
	} else if val_type == "float" {
		f, err := val.Float()
		if err != nil {
			f = 0.0
		}
		buf.WriteString(fmt.Sprintf("%s[%q] = %f,\n", padding(nest_level), key, f))
	} else if val_type == "boolean" {
		b := val.Bool()
		buf.WriteString(fmt.Sprintf("%s[%q] = %v,\n", padding(nest_level), key, b))
	} else {
		panic("invalid atom type")
	}
}

func parse_list(buf *bytes.Buffer, key string, val string, val_type string, nest_level int, nest_list int) {
	val = strings.Trim(val, " ")
	if len(val) == 0 {
		return
	}

	val_list := strings.Split(val, SPLIT[nest_list-1])

	for i := 0; i < len(val_list); i++ {
		if nest_list > 1 {
			buf.WriteString(fmt.Sprintf("%s{\n", padding(nest_level+1)))
			parse_list(buf, key, val_list[i], val_type, nest_level+1, nest_list-1)
			buf.WriteString(fmt.Sprintf("%s},\n", padding(nest_level+1)))
		} else {
			// atom
			if val_type == "string" {
				buf.WriteString(fmt.Sprintf("%s%q,\n", padding(nest_level+1), val_list[i]))
			} else {
				buf.WriteString(fmt.Sprintf("%s%s,\n", padding(nest_level+1), val_list[i]))
			}
		}
	}
}

func padding(nest_level int) string {
	var pad []byte
	indent := []byte("    ")
	for i := 0; i < nest_level; i++ {
		pad = append(pad, indent...)
	}
	return string(pad)
}

func main() {
	if len(os.Args) < 3 {
		fmt.Println("Usage: excel2lua input_filename output_filename")
		return
	}

	input_filename := os.Args[1]
	output_filename := os.Args[2]
	parse(input_filename, output_filename)
}
