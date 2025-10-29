package main

import (
	"bytes"
	"embed"
	"encoding/json"
	"log"
	"net/http"
	"os"
	"time"
        "fmt"

	"github.com/xuri/excelize/v2"
)

// ─────────────────────────────────────────────────────────────────────────────
// Embed the Excel template (keep the file at repo root as template.xlsx)
//
//go:embed template.xlsx
var templateFS embed.FS

// ─────────────────────────────────────────────────────────────────────────────
// Request/Row models (must match your iOS payload)
type Row struct {
	Date    string  `json:"date"`
	Project string  `json:"project"`
	Hours   float64 `json:"hours"`
	Type    string  `json:"type"`
	Notes   string  `json:"notes"`
}

type Request struct {
	EmployeeName string  `json:"employeeName"`
	WeekNumber   int     `json:"weekNumber"` // 1 = Week 1, 2 = Week 2
	Rows         []Row   `json:"rows"`
	TotalOC      float64 `json:"totalOC"`
	TotalOT      float64 `json:"totalOT"`
}

// ─────────────────────────────────────────────────────────────────────────────
// 📍 Layout mapping (match these to your workbook)

// First data row (Sunday)
const startRow = 5

// Employee boxes (✅ set these to your actual cells)
const cellEmployeeW1 = "M2"  // Week 1 Employee cell
const cellEmployeeW2 = "AJ2" // Week 2 Employee cell (adjust if different)

// Column letters for the detail rows.
// Right now we only write the Date (col B). Leave others "" to skip.
type colMap struct{ Date, Project, Hours, Type, Notes string }

var cols = colMap{
	Date:    "B",
	Project: "",
	Hours:   "",
	Type:    "",
	Notes:   "",
}

// Optional: Overtime header "Date:" cell (the single date under the OT table header).
// Leave blank to disable; set to the proper cell if you want the service to fill it.
// Example guesses shown; change if needed.
const cellOTHeaderDateW1 = "" // e.g. "B16"
const cellOTHeaderDateW2 = "" // e.g. "B16"

// Optional: Office Use Only totals on the right side (Regular/OT/DT).
// Leave blank to disable; set to your exact cells to enable.
const (
	cellTotalRegularW1 = "" // e.g. "AJ12"
	cellTotalOTW1      = "" // e.g. "AJ13"
	cellTotalDTW1      = "" // e.g. "AJ14"

	cellTotalRegularW2 = "" // e.g. "AJ31"
	cellTotalOTW2      = "" // e.g. "AJ32"
	cellTotalDTW2      = "" // e.g. "AJ33"
)

// ─────────────────────────────────────────────────────────────────────────────
// Small helpers

// write only if a column/cell is provided
func setIfCol(f *excelize.File, sheet, col string, row int, v any) {
	if col == "" {
		return
	}
	_ = f.SetCellValue(sheet, col+itoa(row), v)
}

// write to an absolute cell (like "M2"), only if provided
func setIfCell(f *excelize.File, sheet, cell string, v any) {
	if cell == "" {
		return
	}
	_ = f.SetCellValue(sheet, cell, v)
}

// itoa without importing strconv
func itoa(i int) string {
	if i == 0 {
		return "0"
	}
	sign := ""
	if i < 0 {
		sign = "-"
		i = -i
	}
	var buf [20]byte
	pos := len(buf)
	for i > 0 {
		pos--
		buf[pos] = byte('0' + i%10)
		i /= 10
	}
	return sign + string(buf[pos:])
}

// ─────────────────────────────────────────────────────────────────────────────
// HTTP handler

func makeHandler(w http.ResponseWriter, r *http.Request) {
	// CORS for local/app usage
	w.Header().Set("Access-Control-Allow-Origin", "*")
	if r.Method == http.MethodOptions {
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")
		w.WriteHeader(http.StatusNoContent)
		return
	}
	if r.Method != http.MethodPost {
		http.Error(w, "POST /excel only", http.StatusMethodNotAllowed)
		return
	}

	var req Request
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "bad json: "+err.Error(), http.StatusBadRequest)
		return
	}

	// Decide which sheet gets the rows (we will still write the name to both)
	sheet := "Week 1"
	if req.WeekNumber == 2 {
		sheet = "Week 2"
	}

	// Open the embedded template
	tmplBytes, err := templateFS.ReadFile("template.xlsx")
	if err != nil {
		http.Error(w, "template read: "+err.Error(), http.StatusInternalServerError)
		return
	}
	f, err := excelize.OpenReader(bytes.NewReader(tmplBytes))
	if err != nil {
		http.Error(w, "open template: "+err.Error(), http.StatusInternalServerError)
		return
	}
	defer f.Close()

	// ✅ Write employee name to BOTH weeks (bypass Excel formula recalc issues)
	setIfCell(f, "Week 1", cellEmployeeW1, req.EmployeeName)
	setIfCell(f, "Week 2", cellEmployeeW2, req.EmployeeName)

	// Detail rows
	rowIdx := startRow
	for _, rr := range req.Rows {
		// Date as Excel date if parseable
		if t, e := time.Parse("2006-01-02", rr.Date); e == nil {
			setIfCol(f, sheet, cols.Date, rowIdx, t)
		} else {
			setIfCol(f, sheet, cols.Date, rowIdx, rr.Date)
		}

		// Others skipped unless you map them to real columns above
		setIfCol(f, sheet, cols.Project, rowIdx, rr.Project)
		setIfCol(f, sheet, cols.Hours, rowIdx, rr.Hours)
		setIfCol(f, sheet, cols.Type, rowIdx, rr.Type)
		setIfCol(f, sheet, cols.Notes, rowIdx, rr.Notes)

		rowIdx++
	}

	// Optional: set OT header "Date:" to the first row's date
	if len(req.Rows) > 0 && req.Rows[0].Date != "" {
		var otDate any = req.Rows[0].Date
		if t, e := time.Parse("2006-01-02", req.Rows[0].Date); e == nil {
			otDate = t
		}
		if sheet == "Week 1" {
			setIfCell(f, "Week 1", cellOTHeaderDateW1, otDate)
		} else {
			setIfCell(f, "Week 2", cellOTHeaderDateW2, otDate)
		}
	}

	// Optional: write right-side totals if you mapped those cells
	writeTotals := func(s string, reg, ot, dt float64) {
		switch s {
		case "Week 1":
			if cellTotalRegularW1 != "" && reg != 0 {
				_ = f.SetCellValue(s, cellTotalRegularW1, reg)
			}
			if cellTotalOTW1 != "" && ot != 0 {
				_ = f.SetCellValue(s, cellTotalOTW1, ot)
			}
			if cellTotalDTW1 != "" && dt != 0 {
				_ = f.SetCellValue(s, cellTotalDTW1, dt)
			}
		case "Week 2":
			if cellTotalRegularW2 != "" && reg != 0 {
				_ = f.SetCellValue(s, cellTotalRegularW2, reg)
			}
			if cellTotalOTW2 != "" && ot != 0 {
				_ = f.SetCellValue(s, cellTotalOTW2, ot)
			}
			if cellTotalDTW2 != "" && dt != 0 {
				_ = f.SetCellValue(s, cellTotalDTW2, dt)
			}
		}
	}
	// Example: we only have TotalOT; Regular/DT set to 0 by default.
	writeTotals(sheet, 0, req.TotalOT, 0)

	// Stream as .xlsx
	buf, err := f.WriteToBuffer()
	if err != nil {
		http.Error(w, "write xlsx: "+err.Error(), http.StatusInternalServerError)
		return
	}
w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
w.Header().Set("Content-Disposition", `attachment; filename="Timecard.xlsx"`)
w.Header().Set("Content-Length", fmt.Sprintf("%d", len(buf.Bytes())))
w.WriteHeader(http.StatusOK)
if _, err := w.Write(buf.Bytes()); err != nil {
    log.Println("write error:", err)

}

// ─────────────────────────────────────────────────────────────────────────────
// main (local + Render)

func main() {
	// health endpoint (useful on Render)
	http.HandleFunc("/health", func(w http.ResponseWriter, _ *http.Request) {
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write([]byte("ok"))
	})

	http.HandleFunc("/excel", makeHandler)

	port := os.Getenv("PORT") // Render sets PORT
	if port == "" {
		port = "8080" // local dev
	}
	log.Println("listening on :" + port)
	log.Fatal(http.ListenAndServe(":"+port, nil))
}
