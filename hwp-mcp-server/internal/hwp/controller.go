package hwp

import (
	"fmt"
	"os"
	"runtime"
	"strings"
	"sync"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/mark3labs/mcp-go/mcp"
)

// Controller wraps the HWP COM interface
type Controller struct {
	hwp         *ole.IDispatch
	visible     bool
	isRunning   bool
	currentPath string
}

var globalController *Controller
var comInitMutex sync.Mutex
var hwpOperationCh chan func()
var hwpOperationOnce sync.Once

func init() {
	globalController = &Controller{}
	// Initialize HWP operation channel for single-threaded COM operations
	initHWPOperationChannel()
}

// GetGlobalController returns the global HWP controller instance
func GetGlobalController() *Controller {
	return globalController
}

// SetGlobalController sets the global HWP controller instance
func SetGlobalController(controller *Controller) {
	globalController = controller
}

// initHWPOperationChannel initializes a single-threaded channel for HWP operations
func initHWPOperationChannel() {
	hwpOperationOnce.Do(func() {
		hwpOperationCh = make(chan func(), 100)
		go func() {
			// Lock this goroutine to a single OS thread for COM operations
			runtime.LockOSThread()
			
			// Initialize COM for this dedicated thread
			ole.CoInitialize(0)
			defer ole.CoUninitialize()
			
			// Process all HWP operations on this single thread
			for operation := range hwpOperationCh {
				operation()
			}
		}()
	})
}

// ExecuteHWPOperation executes a HWP operation on the dedicated COM thread
func ExecuteHWPOperation(operation func()) {
	done := make(chan struct{})
	hwpOperationCh <- func() {
		operation()
		close(done)
	}
	<-done
}

// ExecuteHWPOperationWithResult executes a HWP operation and returns a result
func ExecuteHWPOperationWithResult[T any](operation func() T) T {
	done := make(chan T, 1)
	hwpOperationCh <- func() {
		done <- operation()
	}
	return <-done
}

// ExecuteHWPOperationWithError executes a HWP operation that can return an error
func ExecuteHWPOperationWithError(operation func() error) error {
	return ExecuteHWPOperationWithResult(operation)
}

// ensureCOMInitialized ensures COM is initialized for the current thread
func ensureCOMInitialized() {
	comInitMutex.Lock()
	defer comInitMutex.Unlock()
	
	// Lock OS thread to ensure COM calls happen on the same thread
	runtime.LockOSThread()
	ole.CoInitialize(0)
}

// safeCallMethod safely calls a COM method with panic recovery
func safeCallMethod(obj *ole.IDispatch, method string, params ...interface{}) (result *ole.VARIANT, err error) {
	defer func() {
		if r := recover(); r != nil {
			err = fmt.Errorf("COM method call panic: %v", r)
		}
	}()
	
	if obj == nil {
		return nil, fmt.Errorf("COM object is nil")
	}
	
	result, err = oleutil.CallMethod(obj, method, params...)
	return result, err
}

// safeGetProperty safely gets a COM property with panic recovery
func safeGetProperty(obj *ole.IDispatch, property string) (result *ole.VARIANT, err error) {
	defer func() {
		if r := recover(); r != nil {
			err = fmt.Errorf("COM property access panic: %v", r)
		}
	}()
	
	if obj == nil {
		return nil, fmt.Errorf("COM object is nil")
	}
	
	result, err = oleutil.GetProperty(obj, property)
	return result, err
}

// CreateTextResult creates a text result for MCP responses
func CreateTextResult(text string) *mcp.CallToolResult {
	return &mcp.CallToolResult{
		Content: []mcp.Content{
			mcp.TextContent{
				Type: "text",
				Text: text,
			},
		},
	}
}

// NewController creates a new Controller instance
func NewController() *Controller {
	return &Controller{}
}

// Connect connects to HWP application
func (h *Controller) Connect(visible bool) error {
	// Clean up existing connection if any
	if h.hwp != nil {
		h.hwp.Release()
		h.hwp = nil
	}
	
	unknown, err := oleutil.CreateObject("HWPFrame.HwpObject")
	if err != nil {
		return fmt.Errorf("failed to create HWP object (HWP may not be installed): %v", err)
	}
	// Note: Do NOT defer unknown.Release() here as we need the object to persist

	h.hwp, err = unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		unknown.Release() // Clean up on error
		return fmt.Errorf("failed to query interface: %v", err)
	}
	
	// Store the original unknown object for later cleanup
	// Release the unknown object since we have the IDispatch interface
	unknown.Release()

	// Test if HWP is actually functional by trying to access basic properties
	// Skip version check as it might not be available in all HWP versions
	// Instead, we'll test functionality when actually using methods

	h.visible = visible
	h.isRunning = true

	// Set visibility safely
	if err := h.setVisibility(visible); err != nil {
		// Log error to stderr but don't fail connection
		fmt.Fprintf(os.Stderr, "Warning: Failed to set visibility: %v\n", err)
	}

	return nil
}

// setVisibility sets the HWP window visibility safely
func (h *Controller) setVisibility(visible bool) error {
	// Try to get XHwpWindows property safely
	windowsVar, err := safeGetProperty(h.hwp, "XHwpWindows")
	if err != nil {
		return fmt.Errorf("failed to get XHwpWindows property: %v", err)
	}
	defer windowsVar.Clear()

	windows := windowsVar.ToIDispatch()
	if windows == nil {
		return fmt.Errorf("XHwpWindows is nil")
	}

	// Try to get Item(0) safely
	windowVar, err := safeCallMethod(windows, "Item", 0)
	if err != nil {
		return fmt.Errorf("failed to get window item: %v", err)
	}
	defer windowVar.Clear()

	window := windowVar.ToIDispatch()
	if window == nil {
		return fmt.Errorf("window item is nil")
	}

	// Set visibility safely
	defer func() {
		if r := recover(); r != nil {
			fmt.Fprintf(os.Stderr, "Recovered from panic in PutProperty: %v\n", r)
		}
	}()
	
	if _, err := oleutil.PutProperty(window, "Visible", visible); err != nil {
		return fmt.Errorf("failed to set visibility: %v", err)
	}

	return nil
}

// Disconnect disconnects from HWP application
func (h *Controller) Disconnect() error {
	if h.hwp != nil {
		h.hwp.Release()
		h.hwp = nil
	}
	h.isRunning = false
	h.visible = false
	h.currentPath = ""
	return nil
}

// IsRunning returns whether HWP is running
func (h *Controller) IsRunning() bool {
	return h.isRunning
}

// GetHwp returns the HWP dispatch interface
func (h *Controller) GetHwp() *ole.IDispatch {
	return h.hwp
}

// CreateNewDocument creates a new document
func (h *Controller) CreateNewDocument() error {
	// Always ensure we have a valid connection
	if !h.isRunning || h.hwp == nil {
		if err := h.Connect(true); err != nil {
			return err
		}
	}
	
	// Test if connection is still valid
	if h.hwp == nil {
		return fmt.Errorf("HWP connection is not available")
	}
	
	// Create new document using HAction
	hActionVar, err := safeGetProperty(h.hwp, "HAction")
	if err != nil {
		return fmt.Errorf("failed to get HAction: %v", err)
	}
	defer hActionVar.Clear()
	
	hAction := hActionVar.ToIDispatch()
	if hAction == nil {
		return fmt.Errorf("HAction is nil")
	}
	
	_, err = safeCallMethod(hAction, "Run", "FileNew")
	if err != nil {
		return fmt.Errorf("failed to create new document: %v", err)
	}
	
	h.currentPath = ""
	return nil
}

// OpenDocument opens a document
func (h *Controller) OpenDocument(path string) error {
	if !h.isRunning {
		if err := h.Connect(true); err != nil {
			return err
		}
	}
	
	_, err := safeCallMethod(h.hwp, "Open", path)
	if err == nil {
		h.currentPath = path
	}
	return err
}

// SaveDocument saves the document
func (h *Controller) SaveDocument(path string) error {
	if !h.isRunning || h.hwp == nil {
		return fmt.Errorf("HWP not connected")
	}

	if path != "" {
		_, err := safeCallMethod(h.hwp, "SaveAs", path, "HWP", "")
		if err == nil {
			h.currentPath = path
		}
		return err
	} else if h.currentPath != "" {
		_, err := safeCallMethod(h.hwp, "Save")
		return err
	} else {
		_, err := safeCallMethod(h.hwp, "SaveAs")
		return err
	}
}

// InsertText inserts text at current cursor position
func (h *Controller) InsertText(text string, preserveLinebreaks bool) error {
	if !h.isRunning || h.hwp == nil {
		return fmt.Errorf("HWP not connected")
	}

	if preserveLinebreaks && strings.Contains(text, "\n") {
		lines := strings.Split(text, "\n")
		for i, line := range lines {
			if i > 0 {
				if err := h.InsertParagraph(); err != nil {
					return err
				}
			}
			if strings.TrimSpace(line) != "" {
				if err := h.insertTextDirect(line); err != nil {
					return err
				}
			}
		}
		return nil
	}

	return h.insertTextDirect(text)
}

func (h *Controller) insertTextDirect(text string) error {
	if h.hwp == nil {
		return fmt.Errorf("HWP connection is not available")
	}
	
	// Safely get HAction property
	hActionVar, err := safeGetProperty(h.hwp, "HAction")
	if err != nil {
		return fmt.Errorf("failed to get HAction: %v", err)
	}
	defer hActionVar.Clear()
	
	hAction := hActionVar.ToIDispatch()
	if hAction == nil {
		return fmt.Errorf("HAction is nil")
	}

	// Safely get HParameterSet property
	hParameterSetVar, err := safeGetProperty(h.hwp, "HParameterSet")
	if err != nil {
		return fmt.Errorf("failed to get HParameterSet: %v", err)
	}
	defer hParameterSetVar.Clear()
	
	hParameterSet := hParameterSetVar.ToIDispatch()
	if hParameterSet == nil {
		return fmt.Errorf("HParameterSet is nil")
	}

	// Safely get HInsertText property
	hInsertTextVar, err := safeGetProperty(hParameterSet, "HInsertText")
	if err != nil {
		return fmt.Errorf("failed to get HInsertText: %v", err)
	}
	defer hInsertTextVar.Clear()
	
	hInsertText := hInsertTextVar.ToIDispatch()
	if hInsertText == nil {
		return fmt.Errorf("HInsertText is nil")
	}

	// Safely get HSet property
	hSetVar, err := safeGetProperty(hInsertText, "HSet")
	if err != nil {
		return fmt.Errorf("failed to get HSet: %v", err)
	}
	defer hSetVar.Clear()
	
	hSet := hSetVar.ToIDispatch()
	if hSet == nil {
		return fmt.Errorf("HSet is nil")
	}

	// Execute the text insertion safely
	if _, err := safeCallMethod(hAction, "GetDefault", "InsertText", hSet); err != nil {
		return fmt.Errorf("failed to get default: %v", err)
	}

	// Set text property safely
	defer func() {
		if r := recover(); r != nil {
			fmt.Fprintf(os.Stderr, "Recovered from panic in PutProperty Text: %v\n", r)
		}
	}()
	
	if _, err := oleutil.PutProperty(hInsertText, "Text", text); err != nil {
		return fmt.Errorf("failed to set text property: %v", err)
	}

	if _, err := safeCallMethod(hAction, "Execute", "InsertText", hSet); err != nil {
		return fmt.Errorf("failed to execute insert text: %v", err)
	}

	return nil
}

// SetFontStyle sets font style properties
func (h *Controller) SetFontStyle(fontName string, fontSize int, bold, italic, underline bool) error {
	if !h.isRunning {
		return fmt.Errorf("HWP not connected")
	}

	hAction := oleutil.MustGetProperty(h.hwp, "HAction").ToIDispatch()
	hParameterSet := oleutil.MustGetProperty(h.hwp, "HParameterSet").ToIDispatch()
	hCharShape := oleutil.MustGetProperty(hParameterSet, "HCharShape").ToIDispatch()
	hSet := oleutil.MustGetProperty(hCharShape, "HSet").ToIDispatch()

	oleutil.CallMethod(hAction, "GetDefault", "CharShape", hSet)

	if fontName != "" {
		oleutil.PutProperty(hCharShape, "FaceNameHangul", fontName)
		oleutil.PutProperty(hCharShape, "FaceNameLatin", fontName)
		oleutil.PutProperty(hCharShape, "FaceNameHanja", fontName)
		oleutil.PutProperty(hCharShape, "FaceNameJapanese", fontName)
		oleutil.PutProperty(hCharShape, "FaceNameOther", fontName)
		oleutil.PutProperty(hCharShape, "FaceNameSymbol", fontName)
		oleutil.PutProperty(hCharShape, "FaceNameUser", fontName)
	}

	if fontSize > 0 {
		oleutil.PutProperty(hCharShape, "Height", fontSize*100)
	}

	oleutil.PutProperty(hCharShape, "Bold", bold)
	oleutil.PutProperty(hCharShape, "Italic", italic)
	underlineType := 0
	if underline {
		underlineType = 1
	}
	oleutil.PutProperty(hCharShape, "UnderlineType", underlineType)

	_, err := oleutil.CallMethod(hAction, "Execute", "CharShape", hSet)
	return err
}

// InsertParagraph inserts a new paragraph
func (h *Controller) InsertParagraph() error {
	if !h.isRunning {
		return fmt.Errorf("HWP not connected")
	}

	hAction := oleutil.MustGetProperty(h.hwp, "HAction").ToIDispatch()
	_, err := oleutil.CallMethod(hAction, "Run", "BreakPara")
	return err
}

// GetText gets the document text
func (h *Controller) GetText() (string, error) {
	if !h.isRunning {
		return "", fmt.Errorf("HWP not connected")
	}

	result, err := oleutil.CallMethod(h.hwp, "GetTextFile", "TEXT", "")
	if err != nil {
		return "", err
	}
	return result.ToString(), nil
}

// InsertTable inserts a table
func (h *Controller) InsertTable(rows, cols int) error {
	if !h.isRunning {
		return fmt.Errorf("HWP not connected")
	}

	hAction := oleutil.MustGetProperty(h.hwp, "HAction").ToIDispatch()
	hParameterSet := oleutil.MustGetProperty(h.hwp, "HParameterSet").ToIDispatch()
	hTableCreation := oleutil.MustGetProperty(hParameterSet, "HTableCreation").ToIDispatch()
	hSet := oleutil.MustGetProperty(hTableCreation, "HSet").ToIDispatch()

	oleutil.CallMethod(hAction, "GetDefault", "TableCreate", hSet)
	oleutil.PutProperty(hTableCreation, "Rows", rows)
	oleutil.PutProperty(hTableCreation, "Cols", cols)
	oleutil.PutProperty(hTableCreation, "WidthType", 0)
	oleutil.PutProperty(hTableCreation, "HeightType", 1)
	oleutil.PutProperty(hTableCreation, "WidthValue", 0)
	oleutil.PutProperty(hTableCreation, "HeightValue", 1000)

	// Set column widths
	colWidth := 8000 / cols
	oleutil.CallMethod(hTableCreation, "CreateItemArray", "ColWidth", cols)
	colWidthArray := oleutil.MustGetProperty(hTableCreation, "ColWidth").ToIDispatch()
	for i := 0; i < cols; i++ {
		oleutil.CallMethod(colWidthArray, "SetItem", i, colWidth)
	}

	_, err := oleutil.CallMethod(hAction, "Execute", "TableCreate", hSet)
	return err
}

// FillTableWithData fills table with 2D data
func (h *Controller) FillTableWithData(data [][]string, startRow, startCol int, hasHeader bool) error {
	if !h.isRunning {
		return fmt.Errorf("HWP not connected")
	}

	// Move to table start
	oleutil.CallMethod(h.hwp, "Run", "TableSelCell")
	oleutil.CallMethod(h.hwp, "Run", "TableSelTable")
	oleutil.CallMethod(h.hwp, "Run", "Cancel")
	oleutil.CallMethod(h.hwp, "Run", "TableSelCell")
	oleutil.CallMethod(h.hwp, "Run", "Cancel")

	// Move to start position
	for i := 0; i < startRow-1; i++ {
		oleutil.CallMethod(h.hwp, "Run", "TableLowerCell")
	}
	for i := 0; i < startCol-1; i++ {
		oleutil.CallMethod(h.hwp, "Run", "TableRightCell")
	}

	// Fill data
	for rowIdx, rowData := range data {
		for colIdx, cellValue := range rowData {
			oleutil.CallMethod(h.hwp, "Run", "TableSelCell")
			oleutil.CallMethod(h.hwp, "Run", "Delete")

			if hasHeader && rowIdx == 0 {
				h.SetFontStyle("", 0, true, false, false)
				h.insertTextDirect(cellValue)
				h.SetFontStyle("", 0, false, false, false)
			} else {
				h.insertTextDirect(cellValue)
			}

			if colIdx < len(rowData)-1 {
				oleutil.CallMethod(h.hwp, "Run", "TableRightCell")
			}
		}

		if rowIdx < len(data)-1 {
			for i := 0; i < len(rowData)-1; i++ {
				oleutil.CallMethod(h.hwp, "Run", "TableLeftCell")
			}
			oleutil.CallMethod(h.hwp, "Run", "TableLowerCell")
		}
	}

	// Move cursor out of table
	oleutil.CallMethod(h.hwp, "Run", "TableSelCell")
	oleutil.CallMethod(h.hwp, "Run", "Cancel")
	oleutil.CallMethod(h.hwp, "Run", "MoveDown")

	return nil
}