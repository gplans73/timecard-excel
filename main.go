package main

import (
	"bytes"
	"embed"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

//go:embed template.xlsx
var templateFS embed.FS

type Row struct {
	Date    string  `json:"date"`
	Project string  `json:"project"`
	Hours   float64 `json:"hours"`
	Type    string  `json:"type"`
	Notes   string  `json:"notes"`
}

type Request struct {
	EmployeeName string  `json:"employeeName"`
	WeekNumber   int     `json:"weekNumber"`
	Rows         []Row   `json:"rows"`
	TotalOC      float64 `json:"totalOC"`
	TotalOT      float64 `json:"totalOT"`
}

func parseISO(d string) (time.Time, error) {
	formats := []string{"2006-01-02", "06-01-02", "2006/01/02", "01/02/2006", "02-01-2006", "02/01/2006"}
	for _, f := range formats {
		if t, err := time.ParseInLocation(f, d, time.Local); err == nil {
			return t, nil
		}
	}
	return time.Time{}, fmt.Errorf("bad date: %q", d)
}

func makeHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method == http.MethodOptions {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")
		w.WriteHeader(http.StatusNoContent)
		return
	}
	if r.Method != http.MethodPost {
		http.Error(w, "use POST", http.StatusMethodNotAllowed)
		return
	}

	var req Request
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "bad json: "+err.Error(), http.StatusBadRequest)
		return
	}
	if req.WeekNumber != 1 && req.WeekNumber != 2 {
		req.WeekNumber = 1
	}
	if len(req.Rows) < 7 {
		http.Error(w, "need at least 7 rows (Sun..Sat)", http.StatusBadRequest)
		return
	}

	tmpl, err := templateFS.ReadFile("template.xlsx")
	if err != nil {
		http.Error(w, "template read: "+err.Error(), http.StatusInternalServerError)
		return
	}
	f, err := excelize.OpenReader(bytes.NewReader(tmpl))
	if err != nil {
		http.Error(w, "open xlsx: "+err.Error(), http.StatusInternalServerError)
		return
	}
	defer func() { _ = f.Close() }()

	type weekLayout struct {
		sheet          string
		empCell        string
		mainDatesTop   string
		otDatesTop     string
	}

	layout := map[int]weekLayout{
		1: {sheet: "Week 1", empCell: "M2", mainDatesTop: "B5", otDatesTop: "B16"},
		2: {sheet: "Week 2", empCell: "M2", mainDatesTop: "B5", otDatesTop: "B16"},
	}[req.WeekNumber]

	if req.EmployeeName != "" {
		if err := f.SetCellValue(layout.sheet, layout.empCell, req.EmployeeName); err != nil {
			http.Error(w, "set employee: "+err.Error(), http.StatusInternalServerError)
			return
		}
	}

	dateStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		NumFmt:    14, // short date (locale), keeps borders intact
	})
	if err != nil {
		http.Error(w, "date style: "+err.Error(), http.StatusInternalServerError)
		return
	}

	fillDates := func(top string) error {
		col, row, err := excelize.CellNameToCoordinates(top)
		if err != nil {
			return err
		}
		for i := 0; i < 7; i++ {
			cell, _ := excelize.CoordinatesToCellName(col, row+i)
			dt, err := parseISO(req.Rows[i].Date)
			if err != nil {
				continue
			}
			if err := f.SetCellValue(layout.sheet, cell, dt); err != nil {
				return err
			}
			if err := f.SetCellStyle(layout.sheet, cell, cell, dateStyle); err != nil {
				return err
			}
		}
		return nil
	}

	if err := fillDates(layout.mainDatesTop); err != nil {
		http.Error(w, "main dates: "+err.Error(), http.StatusInternalServerError)
		return
	}
	if err := fillDates(layout.otDatesTop); err != nil {
		http.Error(w, "ot dates: "+err.Error(), http.StatusInternalServerError)
		return
	}

	if t0, err := parseISO(req.Rows[0].Date); err == nil {
		// Set the big "Sun Date Start" box to the Sunday of that week
		sunday := t0.AddDate(0, 0, -int((int(t0.Weekday())+7-0)%7))
		_ = f.SetCellValue(layout.sheet, "B4", sunday)
		_ = f.SetCellStyle(layout.sheet, "B4", "B4", dateStyle)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		http.Error(w, "write xlsx: "+err.Error(), http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", `attachment; filename="Timecard.xlsx"`)
	w.Header().Set("Content-Length", fmt.Sprintf("%d", buf.Len()))
	w.WriteHeader(http.StatusOK)
	if _, err := w.Write(buf.Bytes()); err != nil {
		log.Println("write error:", err)
	}
}

func main() {
	http.HandleFunc("/excel", makeHandler)
	http.HandleFunc("/health", func(w http.ResponseWriter, _ *http.Request) {
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write([]byte("ok"))
	})
	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}
	log.Println("listening on :" + port)
	log.Fatal(http.ListenAndServe(":"+port, nil))
}
