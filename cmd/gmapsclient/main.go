package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

const (
	defaultAPIURL    = "https://gmaps-api.fouiguira.com"
	defaultMaxDepth  = 30
	defaultTimeout   = 300
	pollInterval     = 3 * time.Second
	overallJobWait   = 20 * time.Minute
	settingsFileName = "config.json"
)

type settings struct {
	APIURL string `json:"api_url"`
}

type scrapeRequest struct {
	Keyword  string `json:"keyword"`
	Lang     string `json:"lang"`
	MaxDepth int    `json:"max_depth"`
	Timeout  int    `json:"timeout"`
}

type scrapeResponse struct {
	JobID  string `json:"job_id"`
	Status string `json:"status"`
}

type jobStatusResponse struct {
	JobID       string          `json:"job_id"`
	Status      string          `json:"status"`
	Keyword     string          `json:"keyword"`
	ResultCount int             `json:"result_count"`
	Error       string          `json:"error"`
	Results     json.RawMessage `json:"results"`
}

type ui struct {
	app         fyne.App
	window      fyne.Window
	queryInput  *widget.Entry
	logs        *widget.Entry
	runButton   *widget.Button
	statusLabel *widget.Label
	cfg         settings
}

func main() {
	a := app.New()
	w := a.NewWindow("Google Maps Client")
	w.Resize(fyne.NewSize(820, 560))

	u := &ui{
		app:    a,
		window: w,
		cfg:    loadSettings(),
	}
	if strings.TrimSpace(u.cfg.APIURL) == "" {
		u.cfg.APIURL = defaultAPIURL
	}

	u.queryInput = widget.NewEntry()
	u.queryInput.SetPlaceHolder("Example: dentists in berlin")

	u.logs = widget.NewMultiLineEntry()
	u.logs.Wrapping = fyne.TextWrapWord
	u.logs.SetMinRowsVisible(18)
	u.logs.Disable()

	u.statusLabel = widget.NewLabel("Ready")

	u.runButton = widget.NewButton("Run Query And Export XLSX", func() {
		u.startRun()
	})

	settingsButton := widget.NewButtonWithIcon("Settings", theme.SettingsIcon(), func() {
		u.openSettings()
	})

	top := container.NewHBox(
		widget.NewLabel("Google Maps API Client"),
		layout.NewSpacer(),
		settingsButton,
	)

	content := container.NewVBox(
		widget.NewLabel("Search Query"),
		u.queryInput,
		u.runButton,
		u.statusLabel,
		widget.NewSeparator(),
		widget.NewLabel("Logs"),
		u.logs,
	)

	w.SetContent(container.NewBorder(top, nil, nil, nil, content))
	w.ShowAndRun()
}

func (u *ui) startRun() {
	query := strings.TrimSpace(u.queryInput.Text)
	if query == "" {
		dialog.ShowError(fmt.Errorf("query is required"), u.window)
		return
	}

	u.runOnMain(func() {
		u.runButton.Disable()
		u.statusLabel.SetText("Submitting job...")
		u.logs.SetText("")
	})

	go func() {
		defer u.runOnMain(func() {
			u.runButton.Enable()
		})

		u.logf("Using API URL: %s", u.cfg.APIURL)
		u.logf("Submitting query: %s", query)

		jobID, err := submitJob(u.cfg.APIURL, query)
		if err != nil {
			u.fail(err)
			return
		}

		u.logf("Job submitted: %s", jobID)
		u.setStatus("Job running...")

		result, err := waitForCompletion(u.cfg.APIURL, jobID, func(msg string) {
			u.logf("%s", msg)
		})
		if err != nil {
			u.fail(err)
			return
		}

		rows, err := parseRows(result.Results)
		if err != nil {
			u.fail(fmt.Errorf("failed to parse results: %w", err))
			return
		}

		if len(rows) == 0 {
			u.fail(fmt.Errorf("job completed but no result rows were returned"))
			return
		}

		outPath, err := exportXLSX(query, rows)
		if err != nil {
			u.fail(fmt.Errorf("failed to write xlsx: %w", err))
			return
		}

		u.logf("Done. Result count: %d", result.ResultCount)
		u.logf("XLSX generated: %s", outPath)
		u.setStatus("Completed")

		u.runOnMain(func() {
			dialog.ShowInformation("Export Completed", fmt.Sprintf("Saved file:\n%s", outPath), u.window)
		})
	}()
}

func (u *ui) openSettings() {
	apiEntry := widget.NewEntry()
	apiEntry.SetText(strings.TrimSpace(u.cfg.APIURL))
	apiEntry.SetPlaceHolder("https://gmaps-api.fouiguira.com")

	form := widget.NewForm(
		widget.NewFormItem("API URL", apiEntry),
	)

	d := dialog.NewCustomConfirm("Settings", "Save", "Cancel", form, func(ok bool) {
		if !ok {
			return
		}

		newURL := strings.TrimSpace(apiEntry.Text)
		if newURL == "" {
			dialog.ShowError(fmt.Errorf("api url cannot be empty"), u.window)
			return
		}

		if !strings.HasPrefix(newURL, "http://") && !strings.HasPrefix(newURL, "https://") {
			dialog.ShowError(fmt.Errorf("api url must start with http:// or https://"), u.window)
			return
		}

		u.cfg.APIURL = strings.TrimRight(newURL, "/")
		if err := saveSettings(u.cfg); err != nil {
			dialog.ShowError(err, u.window)
			return
		}

		u.logf("Settings updated. API URL set to %s", u.cfg.APIURL)
	}, u.window)
	d.Resize(fyne.NewSize(560, 160))
	d.Show()
}

func (u *ui) fail(err error) {
	u.logf("Error: %v", err)
	u.setStatus("Failed")
	u.runOnMain(func() {
		dialog.ShowError(err, u.window)
	})
}

func (u *ui) setStatus(s string) {
	u.runOnMain(func() {
		u.statusLabel.SetText(s)
	})
}

func (u *ui) logf(format string, args ...any) {
	line := fmt.Sprintf("[%s] %s\n", time.Now().Format("15:04:05"), fmt.Sprintf(format, args...))
	u.runOnMain(func() {
		u.logs.SetText(u.logs.Text + line)
	})
}

func (u *ui) runOnMain(fn func()) {
	u.app.Driver().RunOnMain(fn)
}

func submitJob(apiURL, query string) (string, error) {
	reqBody := scrapeRequest{
		Keyword:  query,
		Lang:     "en",
		MaxDepth: defaultMaxDepth,
		Timeout:  defaultTimeout,
	}

	body, err := json.Marshal(reqBody)
	if err != nil {
		return "", err
	}

	httpClient := &http.Client{Timeout: 30 * time.Second}
	resp, err := httpClient.Post(strings.TrimRight(apiURL, "/")+"/api/v1/scrape", "application/json", bytes.NewReader(body))
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return "", fmt.Errorf("submit failed with status %d", resp.StatusCode)
	}

	var parsed scrapeResponse
	if err := json.NewDecoder(resp.Body).Decode(&parsed); err != nil {
		return "", err
	}
	if parsed.JobID == "" {
		return "", fmt.Errorf("empty job_id from api")
	}

	return parsed.JobID, nil
}

func waitForCompletion(apiURL, jobID string, logger func(string)) (*jobStatusResponse, error) {
	httpClient := &http.Client{Timeout: 30 * time.Second}
	deadline := time.Now().Add(overallJobWait)

	for time.Now().Before(deadline) {
		jobURL := strings.TrimRight(apiURL, "/") + "/api/v1/jobs/" + jobID
		resp, err := httpClient.Get(jobURL)
		if err != nil {
			logger(fmt.Sprintf("poll error: %v", err))
			time.Sleep(pollInterval)
			continue
		}

		if resp.StatusCode < 200 || resp.StatusCode >= 300 {
			_ = resp.Body.Close()
			logger(fmt.Sprintf("poll status %d", resp.StatusCode))
			time.Sleep(pollInterval)
			continue
		}

		var status jobStatusResponse
		err = json.NewDecoder(resp.Body).Decode(&status)
		_ = resp.Body.Close()
		if err != nil {
			logger(fmt.Sprintf("poll decode error: %v", err))
			time.Sleep(pollInterval)
			continue
		}

		logger(fmt.Sprintf("job status=%s result_count=%d", status.Status, status.ResultCount))

		switch strings.ToLower(status.Status) {
		case "completed":
			return &status, nil
		case "failed":
			if status.Error != "" {
				return nil, fmt.Errorf("job failed: %s", status.Error)
			}
			return nil, fmt.Errorf("job failed")
		}

		time.Sleep(pollInterval)
	}

	return nil, fmt.Errorf("job did not complete in %s", overallJobWait)
}

func parseRows(raw json.RawMessage) ([]map[string]any, error) {
	if len(raw) == 0 || string(raw) == "null" {
		return nil, nil
	}

	var rows []map[string]any
	if err := json.Unmarshal(raw, &rows); err == nil {
		return rows, nil
	}

	var generic []any
	if err := json.Unmarshal(raw, &generic); err != nil {
		return nil, err
	}

	rows = make([]map[string]any, 0, len(generic))
	for _, item := range generic {
		m, ok := item.(map[string]any)
		if ok {
			rows = append(rows, m)
		}
	}

	return rows, nil
}

func exportXLSX(query string, rows []map[string]any) (string, error) {
	if len(rows) == 0 {
		return "", fmt.Errorf("no rows to export")
	}

	headersMap := make(map[string]struct{})
	for _, row := range rows {
		for k := range row {
			headersMap[k] = struct{}{}
		}
	}

	headers := make([]string, 0, len(headersMap))
	for h := range headersMap {
		headers = append(headers, h)
	}
	sort.Strings(headers)

	f := excelize.NewFile()
	const sheet = "Results"
	_ = f.SetSheetName("Sheet1", sheet)

	for c, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(c+1, 1)
		if err := f.SetCellValue(sheet, cell, h); err != nil {
			return "", err
		}
	}

	for r, row := range rows {
		for c, h := range headers {
			cell, _ := excelize.CoordinatesToCellName(c+1, r+2)
			if err := f.SetCellValue(sheet, cell, valueToString(row[h])); err != nil {
				return "", err
			}
		}
	}

	_ = f.SetPanes(sheet, &excelize.Panes{Freeze: true, YSplit: 1})

	outDir := defaultOutputDir()
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		return "", err
	}

	name := sanitizeFileName(query)
	if name == "" {
		name = "results"
	}

	outPath := filepath.Join(outDir, fmt.Sprintf("%s_%s.xlsx", name, time.Now().Format("20060102_150405")))
	if err := f.SaveAs(outPath); err != nil {
		return "", err
	}

	return outPath, nil
}

func valueToString(v any) string {
	switch t := v.(type) {
	case nil:
		return ""
	case string:
		return t
	case float64, float32, int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, bool:
		return fmt.Sprintf("%v", t)
	default:
		b, err := json.Marshal(t)
		if err != nil {
			return fmt.Sprintf("%v", t)
		}
		return string(b)
	}
}

func defaultOutputDir() string {
	home, err := os.UserHomeDir()
	if err != nil {
		return "."
	}

	downloads := filepath.Join(home, "Downloads")
	if st, err := os.Stat(downloads); err == nil && st.IsDir() {
		return downloads
	}

	return home
}

func sanitizeFileName(s string) string {
	s = strings.TrimSpace(strings.ToLower(s))
	s = strings.ReplaceAll(s, " ", "_")
	re := regexp.MustCompile(`[^a-z0-9_\-]+`)
	s = re.ReplaceAllString(s, "")
	if len(s) > 80 {
		s = s[:80]
	}
	return s
}

func settingsPath() (string, error) {
	dir, err := os.UserConfigDir()
	if err != nil {
		return "", err
	}

	p := filepath.Join(dir, "gmaps-client")
	if err := os.MkdirAll(p, 0o755); err != nil {
		return "", err
	}

	return filepath.Join(p, settingsFileName), nil
}

func loadSettings() settings {
	path, err := settingsPath()
	if err != nil {
		return settings{APIURL: defaultAPIURL}
	}

	b, err := os.ReadFile(path)
	if err != nil {
		return settings{APIURL: defaultAPIURL}
	}

	var cfg settings
	if err := json.Unmarshal(b, &cfg); err != nil {
		return settings{APIURL: defaultAPIURL}
	}

	if strings.TrimSpace(cfg.APIURL) == "" {
		cfg.APIURL = defaultAPIURL
	}

	return cfg
}

func saveSettings(cfg settings) error {
	path, err := settingsPath()
	if err != nil {
		return err
	}

	b, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}

	return os.WriteFile(path, b, 0o600)
}
