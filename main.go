package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/olekukonko/tablewriter"
	"github.com/tealeg/xlsx"
)

// Define a data struct to hold the xlsx data
type Record struct {
	ProjectName     string
	Task            string
	KeyNodes        string
	StartDate       string
	StarTime        string
	FinishDate      string
	FinishTime      string
	IPBuildEngineer string
	Designer        string
	ProjectManager  string
	CRQ             string
	Status          string
	Comments        string
}

// Main Function
func main() {
	// Get the date input
	dateIn := bufio.NewReader(os.Stdin)
	fmt.Print("Enter date (format: 02/01/2023): ")
	dateStr, err := dateIn.ReadString('\n')
	if err != nil {
		fmt.Println("Error reading input:", err)
		return
	}

	// Remove the trailing newline or carriage return characters
	dateStr = strings.TrimRight(dateStr, "\r\n")

	// Parse the date string into a time.Time value
	t0, err := time.Parse("02/01/2006", dateStr)
	if err != nil {
		fmt.Println("Error parsing date:", err)
		return
	}

	now := time.Now() // Print out today's date
	fmt.Println("\t\t\t## Today's Date:", now.Local().Day(), now.Local().Month(), now.Local().Year(), "##")
	//subtract 7 days from today
	t1 := t0.AddDate(0, 0, -7)
	println("\t\t\t-- H&L Report First Date:", t1.Day(), "/", t1.Month(), "/", t1.Year(), "--")
	t2 := t1.AddDate(0, 0, 7)
	println("\t\t\t-- H&L Report Last Date:", t2.Day(), "/", t2.Month(), "/", t2.Year(), "--")

	// Create a date filter between two dates
	var dateFilters []string
	for dateFilter := t1; dateFilter.Before(t2); dateFilter = dateFilter.AddDate(0, 0, 1) {
		dateStr := dateFilter.Format("02/01/06")
		dateFilters = append(dateFilters, dateStr)
	}

	// Define constants
	const (
		statusCompleted = "completed"
		statusCancelled = "cancelled"
		excelFile       = "D:\\workstack_live.xlsx"
	)

	// Read Excel data
	records, err := readExcelData(excelFile)
	if err != nil {
		fmt.Println("Error reading Excel file:", err)
		return
	}

	// Filter data and create table
	table := tablewriter.NewWriter(os.Stdout)
	table.SetHeader([]string{"ProjectName", "CRQ", "Task", "StartDate", "Status", "Comments"})
	table.SetRowLine(true)
	countRows := 0
	for _, record := range records {
		if (record.Status == statusCompleted || record.Status == statusCancelled) &&
			contains(dateFilters, record.StartDate) {
			var coloredStatus string
			if record.Status == statusCompleted {
				coloredStatus = colorGreen + record.Status + colorReset
			} else {
				coloredStatus = colorRed + record.Status + colorReset
			}
			table.Append([]string{record.ProjectName, record.CRQ, record.Task, record.StartDate, coloredStatus, record.Comments})
			countRows++
		}
	}
	countCompletedConv := strconv.Itoa(countRows)
	table.SetFooter([]string{"", "", "", "", "Total", countCompletedConv})
	table.Render()

	fmt.Println("Press Enter to continue...")
	fmt.Scanln() // Pauses the screen
	fmt.Println("Program will be closed.")
}

// Function to read data from Excel file and return a slice of Record structs
func readExcelData(excelFile string) ([]Record, error) {
	var records []Record

	// Open excel file
	file, err := xlsx.OpenFile(excelFile)
	if err != nil {
		return nil, err
	}

	// Access the first sheet in the file data structure
	sheet := file.Sheets[0]

	// Loop through the rows to populate slice of Record structs
	for _, row := range sheet.Rows {
		var record Record
		for i, cell := range row.Cells {
			switch i {
			case 0:
				record.ProjectName = cell.String()
			case 1:
				record.Task = cell.String()
			case 2:
				record.KeyNodes = cell.String()
			case 3:
				record.StartDate = cell.String()
			case 4:
				record.StarTime = cell.String()
			case 5:
				record.FinishDate = cell.String()
			case 6:
				record.FinishTime = cell.String()
			case 7:
				record.IPBuildEngineer = cell.String()
			case 8:
				record.Designer = cell.String()
			case 9:
				record.ProjectManager = cell.String()
			case 10:
				record.CRQ = cell.String()
			case 11:
				record.Status = cell.String()
			case 12:
				record.Comments = cell.String()
			}
		}
		records = append(records, record)
	}

	return records, nil
}

// Function to check if a slice of strings contains a string
func contains(slice []string, str string) bool {
	for _, s := range slice {
		if s == str {
			return true
		}
	}
	return false
}

// Constants for colored output
const (
	colorReset = "\033[0m"
	colorGreen = "\033[0;32m"
	colorRed   = "\033[0;31m"
)
