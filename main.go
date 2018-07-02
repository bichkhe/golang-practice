package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strings"
	"os"
	"bufio"
	"strconv"
	"time"
	"math"
)

type DataMonth struct {
	NameMonth string
	DayOfMonth int64
	Number int64

}

type Customer struct {
	Name string
	DoctorName string
	Mobile string
	Address string
	Note string
	DataMap map[int]*DataMonth
}

var customer_list []*Customer


func testFindCustomer(name string, mobile string) (*Customer, bool) {
	var i int = 0
	for i =0 ;i<len(customer_list); i++ {
		if customer_list[i].Mobile== mobile {
			return customer_list[i], true
		}
	}
	return nil, false
}

var columnIndexMap map[string]int




func testWriteHeaderXLSX(xlsx* excelize.File, cellname string){
	index := xlsx.NewSheet(cellname)
	// Set value of a cell.
	xlsx.MergeCell(cellname, "A5", "A6")
	xlsx.SetCellValue(cellname, "A5", "TTS")
	xlsx.MergeCell(cellname, "B5", "B6")
	xlsx.SetCellValue(cellname, "C5", "Tên Bệnh nhân")
	xlsx.MergeCell(cellname, "C5", "C6")
	xlsx.SetCellValue(cellname, "C5", "SĐT")
	xlsx.MergeCell(cellname, "D5", "D6")
	xlsx.SetCellValue(cellname, "D5", "Địa chỉ")

	//Tháng
	xlsx.MergeCell(cellname, "E5", "F5")
	xlsx.SetCellValue(cellname, "E5", "Tháng 1")
	xlsx.SetCellValue(cellname, "E6", "Ngày")
	xlsx.SetCellValue(cellname, "F6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "G5", "H5")
	xlsx.SetCellValue(cellname, "G5", "Tháng 2")
	xlsx.SetCellValue(cellname, "G6", "Ngày")
	xlsx.SetCellValue(cellname, "H6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "I5", "J5")
	xlsx.SetCellValue(cellname, "I5", "Tháng 3")
	xlsx.SetCellValue(cellname, "I6", "Ngày")
	xlsx.SetCellValue(cellname, "J6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "K5", "L5")
	xlsx.SetCellValue(cellname, "K5", "Tháng 4")
	xlsx.SetCellValue(cellname, "K6", "Ngày")
	xlsx.SetCellValue(cellname, "L6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "M5", "N5")
	xlsx.SetCellValue(cellname, "M5", "Tháng 5")
	xlsx.SetCellValue(cellname, "M6", "Ngày")
	xlsx.SetCellValue(cellname, "N6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "O5", "P5")
	xlsx.SetCellValue(cellname, "O5", "Tháng 6")
	xlsx.SetCellValue(cellname, "O6", "Ngày")
	xlsx.SetCellValue(cellname, "P6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "Q5", "R5")
	xlsx.SetCellValue(cellname, "Q5", "Tháng 7")
	xlsx.SetCellValue(cellname, "Q6", "Ngày")
	xlsx.SetCellValue(cellname, "R6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "S5", "T5")
	xlsx.SetCellValue(cellname, "S5", "Tháng 8")
	xlsx.SetCellValue(cellname, "S6", "Ngày")
	xlsx.SetCellValue(cellname, "T6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "U5", "V5")
	xlsx.SetCellValue(cellname, "U5", "Tháng 9")
	xlsx.SetCellValue(cellname, "U6", "Ngày")
	xlsx.SetCellValue(cellname, "V6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "W5", "X5")
	xlsx.SetCellValue(cellname, "W5", "Tháng 10")
	xlsx.SetCellValue(cellname, "W6", "Ngày")
	xlsx.SetCellValue(cellname, "X6", "Số lượng")

	//Tháng
	xlsx.MergeCell(cellname, "Y5", "Z5")
	xlsx.SetCellValue(cellname, "Y5", "Tháng 11")
	xlsx.SetCellValue(cellname, "Y6", "Ngày")
	xlsx.SetCellValue(cellname, "Z6", "Số lượng")
	//Tháng
	xlsx.MergeCell(cellname, "AA5", "AB5")
	xlsx.SetCellValue(cellname, "AA5", "Tháng 12")
	xlsx.SetCellValue(cellname, "AA6", "Ngày")
	xlsx.SetCellValue(cellname, "AB6", "Số lượng")


	xlsx.SetActiveSheet(index)

}
func testSaveFile(xlsx * excelize.File, file string){
	// Save xlsx file by the given path.
	fmt.Print("=====SAVE=======")
	err := xlsx.SaveAs(file)
	if err != nil {
		fmt.Println(err)
	}
}

// fractionOfADay provides function to return the integer values for hour,
// minutes, seconds and nanoseconds that comprised a given fraction of a day.
// values would round to 1 us.
func fractionOfADay(fraction float64) (hours, minutes, seconds, nanoseconds int) {

	const (
		c1us  = 1e3
		c1s   = 1e9
		c1day = 24 * 60 * 60 * c1s
	)

	frac := int64(c1day*fraction + c1us/2)
	nanoseconds = int((frac%c1s)/c1us) * c1us
	frac /= c1s
	seconds = int(frac % 60)
	frac /= 60
	minutes = int(frac % 60)
	hours = int(frac / 60)
	return
}

// shiftJulianToNoon provides function to process julian date to noon.
func shiftJulianToNoon(julianDays, julianFraction float64) (float64, float64) {
	switch {
	case -0.5 < julianFraction && julianFraction < 0.5:
		julianFraction += 0.5
	case julianFraction >= 0.5:
		julianDays++
		julianFraction -= 0.5
	case julianFraction <= -0.5:
		julianDays--
		julianFraction += 1.5
	}
	return julianDays, julianFraction
}
func julianDateToGregorianTime(part1, part2 float64) time.Time {
	part1I, part1F := math.Modf(part1)
	part2I, part2F := math.Modf(part2)
	julianDays := part1I + part2I
	julianFraction := part1F + part2F
	julianDays, julianFraction = shiftJulianToNoon(julianDays, julianFraction)
	day, month, year := doTheFliegelAndVanFlandernAlgorithm(int(julianDays))
	hours, minutes, seconds, nanoseconds := fractionOfADay(julianFraction)
	return time.Date(year, time.Month(month), day, hours, minutes, seconds, nanoseconds, time.UTC)
}

func timeFromExcelTime(excelTime float64, date1904 bool) time.Time {
	var date time.Time
	var intPart = int64(excelTime)
	// Excel uses Julian dates prior to March 1st 1900, and Gregorian
	// thereafter.
	if intPart <= 61 {
		const OFFSET1900 = 15018.0
		const OFFSET1904 = 16480.0
		const MJD0 float64 = 2400000.5
		var date time.Time
		if date1904 {
			date = julianDateToGregorianTime(MJD0, excelTime+OFFSET1904)
		} else {
			date = julianDateToGregorianTime(MJD0, excelTime+OFFSET1900)
		}
		return date
	}
	var floatPart = excelTime - float64(intPart)
	var dayNanoSeconds float64 = 24 * 60 * 60 * 1000 * 1000 * 1000
	if date1904 {
		date = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
	} else {
		date = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	}
	durationDays := time.Duration(intPart) * time.Hour * 24
	durationPart := time.Duration(dayNanoSeconds * floatPart)
	return date.Add(durationDays).Add(durationPart)
}

func doTheFliegelAndVanFlandernAlgorithm(jd int) (day, month, year int) {
	l := jd + 68569
	n := (4 * l) / 146097
	l = l - (146097*n+3)/4
	i := (4000 * (l + 1)) / 1461001
	l = l - (1461*i)/4 + 31
	j := (80 * l) / 2447
	d := l - (2447*j)/80
	l = j / 11
	m := j + 2 - (12 * l)
	y := 100*(n-49) + i + l
	return d, m, y
}
func testReadXLSX(file string, sheetName string, doctor string){
	fmt.Printf ("================== READING %s ====================", file)
	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}

	// Get all the rows in the Sheet1.
	rows := xlsx.GetRows(sheetName)
	for _, row := range rows {
		if strings.Contains(row[3], doctor) == true {
			fmt.Printf("%s -- %s --%s--%s\n", row[1], row[3],row[6], row[15])

			nameArr := strings.Split(row[3], "-")
			addressArr := strings.Split(row[15], "-")
			//fmt.Print(nameArr[0], addressArr[0])

			day, _  :=strconv.ParseFloat(row[1], 64)
			t := timeFromExcelTime(day, false)

			fmt.Printf("%d-%02d-%02dT%02d:%02d:%02d-00:00\n",
				t.Year(), t.Month(), t.Day(),
				t.Hour(), t.Minute(), t.Second())

			imonth := int(t.Month())
			number,_ := strconv.ParseFloat(row[6], 64)
			fmt.Printf("MONTH=%d, NUMBER=%f\n", imonth,number)

			customer, ok := testFindCustomer(nameArr[0], addressArr[0])
			if ok == true {
				//Update data
				fmt.Printf("FOUND %s,%s\n", customer.Name, customer.Mobile)
				data :=customer.DataMap[imonth]
				data.DayOfMonth = int64(t.Day())

				data.Number =  data.Number + int64(number/30.0)
				fmt.Printf("day-number  %d,%d\n", data.DayOfMonth, data.Number)

			}else {
				//Insert customer list a new record
				mobile := addressArr[0]
				var address string
				if len(addressArr) >1 {
					address = addressArr[1]
				}else {
					address  ="Không rõ"
				}

				monthsMap := make (map[int]*DataMonth)
				monthsMap[1] =& DataMonth{
					NameMonth:"Tháng 1",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[2] =& DataMonth{
					NameMonth:"Tháng 2",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[3] =& DataMonth{
					NameMonth:"Tháng 3",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[4] =& DataMonth{
					NameMonth:"Tháng 4",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[5] =& DataMonth{
					NameMonth:"Tháng 5",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[6] =& DataMonth{
					NameMonth:"Tháng 6",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[7] =& DataMonth{
					NameMonth:"Tháng 7",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[8] =& DataMonth{
					NameMonth:"Tháng 8",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[9] =& DataMonth{
					NameMonth:"Tháng 9",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[10] =& DataMonth{
					NameMonth:"Tháng 10",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[11] =& DataMonth{
					NameMonth:"Tháng 11",
					Number: 0,
					DayOfMonth: 0,
				}
				monthsMap[12] =& DataMonth{
					NameMonth:"Tháng 12",
					Number: 0,
					DayOfMonth: 0,
				}


				monthStr := fmt.Sprintf("Tháng %d", int64(t.Month()))
				monthsMap[imonth] = &DataMonth{
					NameMonth:monthStr,
					DayOfMonth: int64(t.Day()),
					Number:int64(number/30.0),
				}
				customer := &Customer{
					Name: nameArr[0],
					DoctorName: "BS Nguyên",
					Address:  address,
					Mobile:mobile,
					DataMap: monthsMap,
				}
				customer_list = append(customer_list, customer)
				fmt.Printf("NEW RECORD: %s-%s-%d-%d\n",nameArr[0],mobile, monthsMap[imonth].DayOfMonth,
					monthsMap[imonth].Number)

			}

		}
		fmt.Println()
	}
}


func testReadOutputXLSX(file string){
	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	fo, err := os.Create("E:\\go_workspace\\src\\bichkhe\\github.com\\bichkhe\\firstproject\\bin\\output.txt")
	if err != nil {
		panic(err)
	}
	// close fo on exit and check for its returned error
	defer func() {
		if err := fo.Close(); err != nil {
			panic(err)
		}
	}()
	// make a write buffer
	w := bufio.NewWriter(fo)

	// Get all the rows in the Sheet1.
	rows := xlsx.GetRows("au2018")
	for idx, row := range rows {
		if(idx >= 9) {
			fmt.Printf("%s--%s--%s--%s--%s--%s\n", row[1], row[2], row[9], row[10], row[11], row[12])

			if len(row[1]) ==0{
				continue
			}

			//WRITE TO FILE
			w.WriteString(row[1])
			w.WriteString(",")
			w.WriteString(row[2])
			w.WriteString(",")

			for  i:=0 ; i<27; i++ {
				if len(row[i])==0 {
					row[i]= "0"
				}
				w.WriteString(row[i])
				w.WriteString(",")

			}

			w.WriteString("\r\n")


			// SAVE INTO MAP
			//months := make([]DataMonth, 0)
			monthsMap := make(map[int]*DataMonth, 0)
			i := 3
			for j:=1; j<=12; j++ {

				dayStr := row[i]
				i++
				numberStr:= row[i]
				i++
				if len(dayStr)==0 {
					dayStr= "0"
				}
				if len(numberStr)==0 {
					numberStr= "0"
				}

				day,_ := strconv.ParseInt(dayStr, 10, 32)
				number,_ := strconv.ParseInt(numberStr, 10, 32)
				monthStr :=fmt.Sprintf("Tháng %d", j)

				//fmt.Printf("==============%s-%s-%s", dayStr, numberStr, monthStr)

				data := &DataMonth{
					NameMonth: monthStr,
					DayOfMonth: day,
					Number: number,
				}
				//months = append(months, data)

				monthsMap[j] = data
			}


			addressArr := strings.Split(row[2], "-")

			mobile := addressArr[0]
			var address string
			if len(addressArr) >1 {
				address = addressArr[1]
			}else {
				address  ="Không rõ"
			}
			customer := &Customer{
				Name: row[1],
				DoctorName: "BS Nguyên",
				Address:  address,
				Mobile:mobile,
				DataMap: monthsMap,

			}
			fmt.Printf("CUSTOMER:%s, Doctor:%s, Address:%s, Mobile:%s\n",
				row[1],"BS Nguyên", address, mobile)



			for  idx, j:=0, 1; idx<len(monthsMap); idx++{
				fmt.Printf("BEFORE===%s-%d-%d\n", monthsMap[j].NameMonth, monthsMap[j].Number,
					monthsMap[j].DayOfMonth)

				//monthsMap[j].DayOfMonth =
				j++
			}

			//for  idx, j:=0, 1; idx<len(monthsMap); idx++{
			//	fmt.Printf("AFTER===%s-%d-%d\n", monthsMap[j].NameMonth, monthsMap[j].Number,
			//		monthsMap[j].DayOfMonth)
			//	j++
			//}

			customer_list = append(customer_list, customer)


			//w.WriteString(",")
			//w.WriteString(row[10])
			//w.WriteString(",")
			//w.WriteString(row[11])
			//w.WriteString(",")
			//w.WriteString(row[12])
			//w.WriteString(",")
			//w.WriteString(row[13])
			//w.WriteString(",")
			//w.WriteString(row[14])

		}
	}
}

func testWriteXLSX(xlsx *excelize.File, f string, cellName string){

	testWriteHeaderXLSX(xlsx, cellName)

	var i int =10

	for idx:=0; idx <len(customer_list); idx++{

		customer := customer_list[idx]
		colName :=fmt.Sprintf("B%d", i)
		colMobile :=fmt.Sprintf("C%d", i)
		colAddress :=fmt.Sprintf("D%d", i)
		colDayMonthJan :=fmt.Sprintf("E%d", i)
		colNumberOfJan :=fmt.Sprintf("F%d", i)
		colDayMonthFeb :=fmt.Sprintf("G%d", i)
		colNumberOfFeb :=fmt.Sprintf("H%d", i)
		colDayMonthMarch :=fmt.Sprintf("I%d", i)
		colNumberOfMarch :=fmt.Sprintf("J%d", i)
		colDayMonthApril :=fmt.Sprintf("K%d", i)
		colNumberOfApril :=fmt.Sprintf("L%d", i)
		colDayMonthMay :=fmt.Sprintf("M%d", i)
		colNumberOfMay :=fmt.Sprintf("N%d", i)
		colDayMonthJune :=fmt.Sprintf("O%d", i)
		colNumberOfJune :=fmt.Sprintf("P%d", i)
		colDayMonthJuly :=fmt.Sprintf("Q%d", i)
		colNumberOfJuly :=fmt.Sprintf("R%d", i)

		fmt.Printf("=====INSETED %s-%s ============ \n", customer.Name,  customer.Mobile)



		for k,v := range customer.DataMap{
			fmt.Printf("%d-%d-%d\n", k, v.DayOfMonth, v.Number)
		}

		xlsx.SetCellValue(cellName,colName, customer.Name)
		xlsx.SetCellValue(cellName,colMobile, customer.Mobile)
		xlsx.SetCellValue(cellName,colAddress, customer.Address)
		if customer.DataMap[1].Number !=0 {
			xlsx.SetCellValue(cellName, colDayMonthJan, customer.DataMap[1].DayOfMonth)
			xlsx.SetCellValue(cellName, colNumberOfJan, customer.DataMap[1].Number)
		}
		if customer.DataMap[2].Number !=0 {
			xlsx.SetCellValue(cellName, colDayMonthFeb, customer.DataMap[2].DayOfMonth)
			xlsx.SetCellValue(cellName, colNumberOfFeb, customer.DataMap[2].Number)
		}

		if customer.DataMap[3].Number !=0 {
			xlsx.SetCellValue(cellName, colDayMonthMarch, customer.DataMap[3].DayOfMonth)
			xlsx.SetCellValue(cellName, colNumberOfMarch, customer.DataMap[3].Number)
		}

		if customer.DataMap[4].Number !=0 {
			xlsx.SetCellValue(cellName, colDayMonthApril, customer.DataMap[4].DayOfMonth)
			xlsx.SetCellValue(cellName, colNumberOfApril, customer.DataMap[4].Number)
		}
		if customer.DataMap[5].Number !=0 {
			xlsx.SetCellValue(cellName, colDayMonthMay, customer.DataMap[5].DayOfMonth)
			xlsx.SetCellValue(cellName, colNumberOfMay, customer.DataMap[5].Number)
		}
		if customer.DataMap[6].Number !=0 {
			xlsx.SetCellValue(cellName, colDayMonthJune, customer.DataMap[6].DayOfMonth)
			xlsx.SetCellValue(cellName, colNumberOfJune, customer.DataMap[6].Number)
		}
		if customer.DataMap[7].Number !=0 {
			xlsx.SetCellValue(cellName, colDayMonthJuly, customer.DataMap[7].DayOfMonth)
			xlsx.SetCellValue(cellName, colNumberOfJuly, customer.DataMap[7].Number)
			
		}


		//for j:=1; j<=12;j++{
		//	xlsx.SetCellValue("LoanTT","", customer.DataMap[j].DayOfMonth)
		//	xlsx.SetCellValue("LoanTT","", customer.DataMap[j].DayOfMonth)
		//}
		i++
	}

	error := xlsx.SaveAs(f)
	if error != nil {
		fmt.Errorf("ERROR%s", error.Error())
	}
}

func main(){

	columnIndexMap = make(map[string]int)
	columnIndexMap["A"] =0
	columnIndexMap["B"] =1
	columnIndexMap["C"] =2
	columnIndexMap["D"] =3
	columnIndexMap["E"] =4
	columnIndexMap["F"] =5
	columnIndexMap["G"] =6
	columnIndexMap["H"] =7
	columnIndexMap["I"] =8
	columnIndexMap["J"] =9
	columnIndexMap["K"] =10
	columnIndexMap["L"] =11
	columnIndexMap["M"] =12
	columnIndexMap["N"] =13
	columnIndexMap["O"] =14
	columnIndexMap["P"] =15
	columnIndexMap["Q"] =16
	columnIndexMap["R"] =17
	columnIndexMap["S"] =18
	columnIndexMap["T"] =19
	columnIndexMap["U"] =20
	columnIndexMap["V"] =21
	columnIndexMap["W"] =22
	columnIndexMap["X"] =23
	columnIndexMap["Y"] =24
	columnIndexMap["Z"] =25





	//testWriteHeaderXLSX()

	//testReadOutputXLSX("xlsx/TDBN_Nguyen_2018.xlsx")

	xlsx := excelize.NewFile()
	doctocs  :=  []string{
		"BS Nguyên",
		"BS Cửu",
		"BS Bình",
		"BS Hưng",
		"BS Hùng",
		"BS Nhật",
		"BS Hương",
		"BS Phúc",
		"BS Thư",
		"BS Trung",
		"BS Điểu",
		"BS Văn Tuấn",
		"BS Huy",
	}

	//var doctor string = "BS Nguyên"

	for _, doctor := range doctocs {
		testReadXLSX("xlsx/TDBN_Thang1.xlsx", "Sheet1", doctor)
		testReadXLSX("xlsx/TDBN_Thang2.xlsx", "Sheet1", doctor)
		testReadXLSX("xlsx/TDBN_Thang3.xlsx", "Sheet1", doctor)
		testReadXLSX("xlsx/TDBN_Thang4.xlsx", "Sheet1", doctor)
		testReadXLSX("xlsx/TDBN_Thang5.xlsx", "Sheet1", doctor)
		testReadXLSX("xlsx/TDBN_Thang6.xlsx", "Sheet1", doctor)

		output := fmt.Sprintf("xlsx/%s.xlsx", "LoanTT")
		testWriteXLSX(xlsx, output, doctor)
		customer_list = customer_list[:0]

	}
	//for  idx, j:=0, 1; idx<len(customer_list); idx++{
	//
	//}
	fmt.Print("======FINISHED ===========")


}
