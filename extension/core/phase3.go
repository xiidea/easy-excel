package core

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"image"
	"os"

	// register decoders for image dimension probing (AddImage scaling)
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"

	"github.com/xuri/excelize/v2"

	"github.com/xiidea/easy-excel/extension/compat"
)

// Phase-3 ops (PLAN.md §5/§13). All of these live in worksheet XML parts
// excelize cannot stream, so they ride the same save-time pending queue as
// auto-filter (structure.go); on never-streamed workbooks they apply at the
// free flag-flip instead.

// SetDataValidation queues a data-validation rule for a range (JSON spec,
// PhpSpreadsheet DataValidation shape).
func (w *Workbook) SetDataValidation(sheet, ref, jsonSpec string) error {
	// validate eagerly: bad specs should fail at the call site
	if _, err := compat.TranslateValidation(ref, jsonSpec); err != nil {
		return err
	}
	return w.queueOp(sheet, pendingOp{kind: opValidation, ref: ref, s1: jsonSpec})
}

// SetConditionalFormat queues conditional-formatting rules for a range
// (JSON list, see compat.TranslateConditionals).
func (w *Workbook) SetConditionalFormat(sheet, ref, jsonRules string) error {
	if _, err := compat.TranslateConditionals(jsonRules); err != nil {
		return err
	}
	return w.queueOp(sheet, pendingOp{kind: opConditional, ref: ref, s1: jsonRules})
}

// imageSpec mirrors the shim's Drawing properties. Either Path (file
// drawing) or Data+Extension (MemoryDrawing, base64 bytes) is set.
type imageSpec struct {
	Path      string `json:"path"`
	Data      string `json:"data"`      // base64-encoded image bytes
	Extension string `json:"extension"` // ".png" / ".jpeg" / ".gif" for Data
	Name      string `json:"name"`
	OffsetX   int    `json:"offsetX"`
	OffsetY   int    `json:"offsetY"`
	Width     int    `json:"width"`  // desired px; 0 = natural size
	Height    int    `json:"height"` // desired px; 0 = natural size
}

// AddImage queues a picture anchored at cell. The file is resolved against
// the path policy now (fail fast) but read by excelize at apply time.
func (w *Workbook) AddImage(sheet, cell, jsonSpec string) error {
	var spec imageSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return fmt.Errorf("easy-excel: invalid image spec: %w", err)
	}
	abs, err := w.policy.Resolve(spec.Path)
	if err != nil {
		return err
	}
	if _, err := os.Stat(abs); err != nil {
		return fmt.Errorf("easy-excel: image: %w", err)
	}
	spec.Path = abs
	encoded, err := json.Marshal(spec)
	if err != nil {
		return err
	}
	return w.queueOp(sheet, pendingOp{kind: opImage, ref: cell, s1: string(encoded)})
}

// protectSpec mirrors PhpSpreadsheet's Worksheet\Protection flags. The
// boolean flags use PhpSpreadsheet's polarity: true = the action is LOCKED
// (disallowed); excelize wants true = allowed, so they are inverted here.
type protectSpec struct {
	Password            string `json:"password"`
	Sheet               bool   `json:"sheet"`
	AutoFilter          bool   `json:"autoFilter"`
	DeleteColumns       bool   `json:"deleteColumns"`
	DeleteRows          bool   `json:"deleteRows"`
	FormatCells         bool   `json:"formatCells"`
	FormatColumns       bool   `json:"formatColumns"`
	FormatRows          bool   `json:"formatRows"`
	InsertColumns       bool   `json:"insertColumns"`
	InsertHyperlinks    bool   `json:"insertHyperlinks"`
	InsertRows          bool   `json:"insertRows"`
	Objects             bool   `json:"objects"`
	PivotTables         bool   `json:"pivotTables"`
	Scenarios           bool   `json:"scenarios"`
	SelectLockedCells   bool   `json:"selectLockedCells"`
	SelectUnlockedCells bool   `json:"selectUnlockedCells"`
	Sort                bool   `json:"sort"`
}

// ProtectSheet queues sheet protection ("sheet": false clears it is not
// supported — PhpSpreadsheet parity: protection applies when sheet=true).
func (w *Workbook) ProtectSheet(sheet, jsonSpec string) error {
	var spec protectSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return fmt.Errorf("easy-excel: invalid protection spec: %w", err)
	}
	if !spec.Sheet {
		return nil // protection not enabled; nothing to queue
	}
	return w.queueOp(sheet, pendingOp{kind: opProtect, ref: "", s1: jsonSpec})
}

// AddChart queues a chart (easy-excel native JSON spec) anchored at cell.
func (w *Workbook) AddChart(sheet, cell, jsonSpec string) error {
	if _, err := compat.TranslateChart(jsonSpec); err != nil {
		return err
	}
	return w.queueOp(sheet, pendingOp{kind: opChart, ref: cell, s1: jsonSpec})
}

// applyOpPhase3 executes the queued Phase-3 ops in random-access mode.
func (w *Workbook) applyOpPhase3(sheet string, op pendingOp) error {
	switch op.kind {
	case opValidation:
		dv, err := compat.TranslateValidation(op.ref, op.s1)
		if err != nil || dv == nil {
			return err
		}
		return w.f.AddDataValidation(sheet, dv)
	case opConditional:
		rules, err := compat.TranslateConditionals(op.s1)
		if err != nil {
			return err
		}
		opts := make([]excelize.ConditionalFormatOptions, len(rules))
		for i, r := range rules {
			opts[i] = r.Options
			if r.Style != nil {
				id, err := w.f.NewConditionalStyle(r.Style)
				if err != nil {
					return err
				}
				opts[i].Format = &id
			}
		}
		return w.f.SetConditionalFormat(sheet, op.ref, opts)
	case opImage:
		return w.applyImage(sheet, op)
	case opProtect:
		var spec protectSpec
		if err := json.Unmarshal([]byte(op.s1), &spec); err != nil {
			return err
		}
		return w.f.ProtectSheet(sheet, &excelize.SheetProtectionOptions{
			Password:            spec.Password,
			AutoFilter:          !spec.AutoFilter,
			DeleteColumns:       !spec.DeleteColumns,
			DeleteRows:          !spec.DeleteRows,
			FormatCells:         !spec.FormatCells,
			FormatColumns:       !spec.FormatColumns,
			FormatRows:          !spec.FormatRows,
			InsertColumns:       !spec.InsertColumns,
			InsertHyperlinks:    !spec.InsertHyperlinks,
			InsertRows:          !spec.InsertRows,
			EditObjects:         !spec.Objects,
			PivotTables:         !spec.PivotTables,
			EditScenarios:       !spec.Scenarios,
			SelectLockedCells:   !spec.SelectLockedCells,
			SelectUnlockedCells: !spec.SelectUnlockedCells,
			Sort:                !spec.Sort,
		})
	case opChart:
		chart, err := compat.TranslateChart(op.s1)
		if err != nil {
			return err
		}
		return w.f.AddChart(sheet, op.ref, chart)
	case opUnmerge:
		tl, br, err := splitRange(op.ref)
		if err != nil {
			return err
		}
		return w.f.UnmergeCell(sheet, tl, br)
	}
	return w.applyOpPhase43(sheet, op)
}

func (w *Workbook) applyImage(sheet string, op pendingOp) error {
	var spec imageSpec
	if err := json.Unmarshal([]byte(op.s1), &spec); err != nil {
		return err
	}
	opts := &excelize.GraphicOptions{
		Name:    spec.Name,
		OffsetX: spec.OffsetX,
		OffsetY: spec.OffsetY,
		ScaleX:  1,
		ScaleY:  1,
		// oneCell anchoring matches PhpSpreadsheet: the image keeps its size
		// when rows/columns resize (twoCell, excelize's default, stretches it)
		Positioning: "oneCell",
	}
	// in-memory drawing: base64 bytes instead of a file path
	if spec.Data != "" {
		raw, err := base64.StdEncoding.DecodeString(spec.Data)
		if err != nil {
			return fmt.Errorf("easy-excel: image data: %w", err)
		}
		if spec.Width > 0 || spec.Height > 0 {
			if cfg, _, err := image.DecodeConfig(bytes.NewReader(raw)); err == nil {
				scaleToFit(opts, spec, cfg.Width, cfg.Height)
			}
		}
		return w.f.AddPictureFromBytes(sheet, op.ref, &excelize.Picture{
			Extension: spec.Extension,
			File:      raw,
			Format:    opts,
		})
	}
	if spec.Width > 0 || spec.Height > 0 {
		fh, err := os.Open(spec.Path)
		if err != nil {
			return err
		}
		cfg, _, err := image.DecodeConfig(fh)
		fh.Close()
		if err != nil {
			return fmt.Errorf("easy-excel: image %s: %w", spec.Path, err)
		}
		scaleToFit(opts, spec, cfg.Width, cfg.Height)
	}
	return w.f.AddPicture(sheet, op.ref, spec.Path, opts)
}

// scaleToFit sets opts.ScaleX/Y so a natWidth×natHeight image renders at the
// spec's requested px; aspect ratio is kept when only one side is given.
func scaleToFit(opts *excelize.GraphicOptions, spec imageSpec, natWidth, natHeight int) {
	if spec.Width > 0 && natWidth > 0 {
		opts.ScaleX = float64(spec.Width) / float64(natWidth)
	}
	if spec.Height > 0 && natHeight > 0 {
		opts.ScaleY = float64(spec.Height) / float64(natHeight)
	}
	if spec.Width > 0 && spec.Height == 0 {
		opts.ScaleY = opts.ScaleX
	}
	if spec.Height > 0 && spec.Width == 0 {
		opts.ScaleX = opts.ScaleY
	}
}
