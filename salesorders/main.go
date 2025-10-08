package main

import (
	"fmt"
	"time"

	"github.com/mehshan/xlsx2struct"
	xlsx3 "github.com/tealeg/xlsx/v3"
)

type SaleOrder struct {
	Date   time.Time `column:"heading=Order Date"`
	Region string    `column:"heading=Region,trim"`
	Rep    string    `column:"heading=Rep"`
	Item   string    `column:"heading=Item,default=Pencil"`
	Units  int32     `column:"heading=Units,default=1"`
	Cost   float32   `column:"heading=Unit Cost"`
	Total  float64   `column:"heading=Total"`
}

func main() {
	orders := []*SaleOrder{}

	file, err := xlsx3.OpenFile("testdata/salesorders.xlsx")
	if err != nil {
		panic(err)
	}

	sheet := file.Sheet["Sales Orders"]
	opt := xlsx2struct.DefaultSheetOptions()

	err = xlsx2struct.Unmarshal(sheet, &orders, opt)
	if err != nil {
		panic(err)
	}

	for _, o := range orders {
		fmt.Printf("%v %s %s %s %d %f %f\n", o.Date, o.Region, o.Rep, o.Item, o.Units, o.Cost, o.Total)
	}
}
