package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
)

// HwpController wraps the HWP COM interface
type HwpController struct {
	hwp         *ole.IDispatch
	visible     bool
	isRunning   bool
	currentPath string
}

// Tool names
const (
	HWP_CREATE               = "hwp_create"
	HWP_OPEN                 = "hwp_open"
	HWP_SAVE                 = "hwp_save"
	HWP_INSERT_TEXT          = "hwp_insert_text"
	HWP_SET_FONT             = "hwp_set_font"
	HWP_INSERT_TABLE         = "hwp_insert_table"
	HWP_INSERT_PARAGRAPH     = "hwp_insert_paragraph"
	HWP_GET_TEXT             = "hwp_get_text"
	HWP_CLOSE                = "hwp_close"
	HWP_PING_PONG            = "hwp_ping_pong"
	HWP_CREATE_TABLE_WITH_DATA = "hwp_create_table_with_data"
	HWP_FILL_TABLE_WITH_DATA = "hwp_fill_table_with_data"
	HWP_FILL_COLUMN_NUMBERS  = "hwp_fill_column_numbers"
	HWP_BATCH_OPERATIONS     = "hwp_batch_operations"
	HWP_CREATE_DOCUMENT_FROM_TEXT = "hwp_create_document_from_text"
	HWP_CREATE_COMPLETE_DOCUMENT = "hwp_create_complete_document"
)

// 채팅 요청 구조체
type ChatRequest struct {
    Messages    []Message `json:"messages"`
    MaxTokens   int       `json:"max_tokens"`
    Temperature float64   `json:"temperature"`
    ModelName   string    `json:"model_name"`
}

type ChatResponse struct {
    Response    string      `json:"response"`
    Status      string      `json:"status"`
    ToolCalls   []ToolCall  `json:"tool_calls,omitempty"`
    NeedsTools  bool        `json:"needs_tools,omitempty"`
}

type ToolCall struct {
    Name      string                 `json:"name"`
    Arguments map[string]interface{} `json:"arguments"`
}

// 메인 채팅 처리 함수
func (app *App) ProcessChatWithMCP(ctx context.Context, userMessage string) (string, error) {
    // 1. LLM 서버에 초기 요청
    chatReq := ChatRequest{
        Messages: []Message{
            {Role: "user", Content: userMessage},
        },
        MaxTokens:   1000,
        Temperature: 0.7,
        ModelName:   "gemini-1.5-flash",
    }
    
    // MCP 도구 정보를 LLM에 제공
    availableTools := app.mcpClient.GetAvailableTools()
    chatReq.AvailableTools = availableTools
    
    response, err := app.sendToLLMServer(chatReq)
    if err != nil {
        return "", fmt.Errorf("LLM server error: %v", err)
    }
    
    // 2. 도구 호출이 필요한지 확인
    if !response.NeedsTools || len(response.ToolCalls) == 0 {
        return response.Response, nil
    }
    
    // 3. MCP 도구들 실행
    toolResults := make([]map[string]interface{}, 0)
    for _, toolCall := range response.ToolCalls {
        result, err := app.mcpClient.CallTool(toolCall.Name, toolCall.Arguments)
        if err != nil {
            toolResults = append(toolResults, map[string]interface{}{
                "tool_name": toolCall.Name,
                "error":     err.Error(),
            })
        } else {
            toolResults = append(toolResults, map[string]interface{}{
                "tool_name": toolCall.Name,
                "result":    result,
            })
        }
    }
    
    // 4. 도구 실행 결과와 함께 최종 응답 요청
    finalChatReq := ChatRequest{
        Messages: []Message{
            {Role: "user", Content: userMessage},
            {Role: "assistant", Content: response.Response},
            {Role: "system", Content: fmt.Sprintf("도구 실행 결과: %v", toolResults)},
        },
        MaxTokens:   1000,
        Temperature: 0.7,
        ModelName:   "gemini-1.5-flash",
    }
    
    finalResponse, err := app.sendToLLMServer(finalChatReq)
    if err != nil {
        return "", fmt.Errorf("final LLM server error: %v", err)
    }
    
    return finalResponse.Response, nil
}

// 연결 해제
func (c *MCPClient) Disconnect() error {
    if c.cmd != nil && c.cmd.Process != nil {
        c.stdin.Close()
        c.stdout.Close()
        c.cmd.Process.Kill()
        c.cmd.Wait()
    }
    c.isConnected = false
    return nil
}

var hwpController *HwpController
var comInitMutex sync.Mutex
var hwpOperationCh chan func()
var hwpOperationOnce sync.Once

func init() {
	hwpController = &HwpController{}
	// Initialize HWP operation channel for single-threaded COM operations
	initHWPOperationChannel()
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

// executeHWPOperation executes a HWP operation on the dedicated COM thread
func executeHWPOperation(operation func()) {
	done := make(chan struct{})
	hwpOperationCh <- func() {
		operation()
		close(done)
	}
	<-done
}

// executeHWPOperationWithResult executes a HWP operation and returns a result
func executeHWPOperationWithResult[T any](operation func() T) T {
	done := make(chan T, 1)
	hwpOperationCh <- func() {
		done <- operation()
	}
	return <-done
}

// executeHWPOperationWithError executes a HWP operation that can return an error
func executeHWPOperationWithError(operation func() error) error {
	return executeHWPOperationWithResult(operation)
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

// Helper function to create text result
func createTextResult(text string) *mcp.CallToolResult {
	return &mcp.CallToolResult{
		Content: []mcp.Content{
			mcp.TextContent{
				Type: "text",
				Text: text,
			},
		},
	}
}

// NewHwpController creates a new HwpController instance
func NewHwpController() *HwpController {
	return &HwpController{}
}

// Connect connects to HWP application
func (h *HwpController) Connect(visible bool) error {
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
func (h *HwpController) setVisibility(visible bool) error {
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
func (h *HwpController) Disconnect() error {
	if h.hwp != nil {
		h.hwp.Release()
		h.hwp = nil
	}
	h.isRunning = false
	h.visible = false
	h.currentPath = ""
	return nil
}

// CreateNewDocument creates a new document
func (h *HwpController) CreateNewDocument() error {
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
func (h *HwpController) OpenDocument(path string) error {
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
func (h *HwpController) SaveDocument(path string) error {
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
func (h *HwpController) InsertText(text string, preserveLinebreaks bool) error {
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

func (h *HwpController) insertTextDirect(text string) error {
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
func (h *HwpController) SetFontStyle(fontName string, fontSize int, bold, italic, underline bool) error {
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

// InsertTable inserts a table
func (h *HwpController) InsertTable(rows, cols int) error {
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

// InsertParagraph inserts a new paragraph
func (h *HwpController) InsertParagraph() error {
	if !h.isRunning {
		return fmt.Errorf("HWP not connected")
	}

	hAction := oleutil.MustGetProperty(h.hwp, "HAction").ToIDispatch()
	_, err := oleutil.CallMethod(hAction, "Run", "BreakPara")
	return err
}

// GetText gets the document text
func (h *HwpController) GetText() (string, error) {
	if !h.isRunning {
		return "", fmt.Errorf("HWP not connected")
	}

	result, err := oleutil.CallMethod(h.hwp, "GetTextFile", "TEXT", "")
	if err != nil {
		return "", err
	}
	return result.ToString(), nil
}

// FillTableWithData fills table with 2D data
func (h *HwpController) FillTableWithData(data [][]string, startRow, startCol int, hasHeader bool) error {
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

// Tool handlers
func handleHwpCreate(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil {
			hwpController = NewHwpController()
		}

		err := hwpController.CreateNewDocument()
		if err != nil {
			// Reset controller on error
			hwpController = nil
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = createTextResult("New document created successfully")
	})
	
	return result, nil
}

func handleHwpOpen(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	path := request.GetString("path", "")
	if path == "" {
		return createTextResult("Error: File path is required"), nil
	}

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil {
			hwpController = NewHwpController()
		}

		err := hwpController.OpenDocument(path)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = createTextResult(fmt.Sprintf("Document opened: %s", path))
	})
	
	return result, nil
}

func handleHwpSave(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	path := request.GetString("path", "")

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := hwpController.SaveDocument(path)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		if path != "" {
			result = createTextResult(fmt.Sprintf("Document saved to: %s", path))
		} else {
			result = createTextResult("Document saved successfully")
		}
	})
	
	return result, nil
}

func handleHwpInsertText(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	text := request.GetString("text", "")
	if text == "" {
		return createTextResult("Error: Text is required"), nil
	}

	preserveLinebreaks := request.GetBool("preserve_linebreaks", true)
	
	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := hwpController.InsertText(text, preserveLinebreaks)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = createTextResult("Text inserted successfully")
	})
	
	return result, nil
}

func handleHwpSetFont(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	name := request.GetString("name", "")
	size := request.GetInt("size", 0)
	bold := request.GetBool("bold", false)
	italic := request.GetBool("italic", false)
	underline := request.GetBool("underline", false)

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := hwpController.SetFontStyle(name, size, bold, italic, underline)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = createTextResult("Font set successfully")
	})
	
	return result, nil
}

func handleHwpInsertTable(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	rows := request.GetInt("rows", 0)
	cols := request.GetInt("cols", 0)

	if rows <= 0 || cols <= 0 {
		return createTextResult("Error: Valid rows and cols are required"), nil
	}
	
	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := hwpController.InsertTable(rows, cols)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = createTextResult(fmt.Sprintf("Table created (%dx%d)", rows, cols))
	})
	
	return result, nil
}

func handleHwpInsertParagraph(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := hwpController.InsertParagraph()
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = createTextResult("Paragraph inserted successfully")
	})
	
	return result, nil
}

func handleHwpGetText(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		text, err := hwpController.GetText()
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = createTextResult(text)
	})
	
	return result, nil
}

func handleHwpClose(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil {
			result = createTextResult("HWP is already closed")
			return
		}

		err := hwpController.Disconnect()
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		hwpController = nil
		result = createTextResult("HWP connection closed successfully")
	})
	
	return result, nil
}

func handleHwpPingPong(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	message := request.GetString("message", "핑")

	var response string
	switch message {
	case "핑":
		response = "퐁"
	case "퐁":
		response = "핑"
	default:
		response = fmt.Sprintf("모르는 메시지입니다: %s (핑 또는 퐁을 보내주세요)", message)
	}

	result := map[string]interface{}{
		"response":         response,
		"original_message": message,
		"timestamp":        time.Now().Format("2006-01-02 15:04:05"),
	}

	resultJSON, _ := json.Marshal(result)
	return createTextResult(string(resultJSON)), nil
}

func handleHwpFillTableWithData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	dataStr := request.GetString("data", "")
	if dataStr == "" {
		return createTextResult("Error: Data is required"), nil
	}

	startRow := request.GetInt("start_row", 1)
	startCol := request.GetInt("start_col", 1)
	hasHeader := request.GetBool("has_header", false)

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Parse JSON data
		var tableData [][]string
		var jsonData [][]interface{}
		if err := json.Unmarshal([]byte(dataStr), &jsonData); err != nil {
			result = createTextResult(fmt.Sprintf("Error: Failed to parse JSON data - %v", err))
			return
		}

		for _, rowInterface := range jsonData {
			var row []string
			for _, cell := range rowInterface {
				row = append(row, fmt.Sprintf("%v", cell))
			}
			tableData = append(tableData, row)
		}

		err := hwpController.FillTableWithData(tableData, startRow, startCol, hasHeader)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = createTextResult("Table data filled successfully")
	})
	
	return result, nil
}

func handleHwpFillColumnNumbers(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	start := request.GetInt("start", 1)
	end := request.GetInt("end", 10)
	column := request.GetInt("column", 1)

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Move to table beginning
		oleutil.CallMethod(hwpController.hwp, "Run", "TableColBegin")

		// Move to specified column
		for i := 0; i < column-1; i++ {
			oleutil.CallMethod(hwpController.hwp, "Run", "TableRightCell")
		}

		// Fill numbers
		for num := start; num <= end; num++ {
			oleutil.CallMethod(hwpController.hwp, "Run", "Select")
			oleutil.CallMethod(hwpController.hwp, "Run", "Delete")
			
			hwpController.insertTextDirect(strconv.Itoa(num))

			if num < end {
				oleutil.CallMethod(hwpController.hwp, "Run", "TableLowerCell")
			}
		}

		result = createTextResult(fmt.Sprintf("Column %d filled with numbers %d~%d", column, start, end))
	})
	
	return result, nil
}

func handleHwpCreateTableWithData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	rows := request.GetInt("rows", 0)
	cols := request.GetInt("cols", 0)
	dataStr := request.GetString("data", "")
	hasHeader := request.GetBool("has_header", false)

	if rows <= 0 || cols <= 0 {
		return createTextResult("Error: Valid rows and cols are required"), nil
	}

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Create table first
		err := hwpController.InsertTable(rows, cols)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error creating table: %v", err))
			return
		}

		// Fill with data if provided
		if dataStr != "" {
			var tableData [][]string
			var jsonData [][]interface{}
			if err := json.Unmarshal([]byte(dataStr), &jsonData); err != nil {
				result = createTextResult(fmt.Sprintf("Error: Failed to parse JSON data - %v", err))
				return
			}

			for _, rowInterface := range jsonData {
				var row []string
				for _, cell := range rowInterface {
					row = append(row, fmt.Sprintf("%v", cell))
				}
				tableData = append(tableData, row)
			}

			err = hwpController.FillTableWithData(tableData, 1, 1, hasHeader)
			if err != nil {
				result = createTextResult(fmt.Sprintf("Error filling table: %v", err))
				return
			}
		}

		result = createTextResult(fmt.Sprintf("Table created (%dx%d) and filled with data", rows, cols))
	})
	
	return result, nil
}

func handleHwpBatchOperations(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	operationsStr := request.GetString("operations", "")
	if operationsStr == "" {
		return createTextResult("Error: Operations list is required"), nil
	}

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil || !hwpController.isRunning || hwpController.hwp == nil {
			result = createTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		var operations []map[string]interface{}
		if err := json.Unmarshal([]byte(operationsStr), &operations); err != nil {
			result = createTextResult(fmt.Sprintf("Error: Failed to parse operations JSON - %v", err))
			return
		}

		var results []string
		for i, op := range operations {
			opType, ok := op["type"].(string)
			if !ok {
				results = append(results, fmt.Sprintf("Operation %d: Error - missing type", i+1))
				continue
			}

			var err error
			switch opType {
			case "insert_text":
				text, _ := op["text"].(string)
				preserveLinebreaks, _ := op["preserve_linebreaks"].(bool)
				err = hwpController.InsertText(text, preserveLinebreaks)
			case "insert_paragraph":
				err = hwpController.InsertParagraph()
			case "set_font":
				name, _ := op["name"].(string)
				size := int(op["size"].(float64))
				bold, _ := op["bold"].(bool)
				italic, _ := op["italic"].(bool)
				underline, _ := op["underline"].(bool)
				err = hwpController.SetFontStyle(name, size, bold, italic, underline)
			case "insert_table":
				rows := int(op["rows"].(float64))
				cols := int(op["cols"].(float64))
				err = hwpController.InsertTable(rows, cols)
			default:
				err = fmt.Errorf("unknown operation type: %s", opType)
			}

			if err != nil {
				results = append(results, fmt.Sprintf("Operation %d (%s): Error - %v", i+1, opType, err))
			} else {
				results = append(results, fmt.Sprintf("Operation %d (%s): Success", i+1, opType))
			}
		}

		resultJSON, _ := json.Marshal(map[string]interface{}{
			"total_operations": len(operations),
			"results": results,
		})
		result = createTextResult(string(resultJSON))
	})
	
	return result, nil
}

func handleHwpCreateDocumentFromText(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	content := request.GetString("content", "")
	if content == "" {
		return createTextResult("Error: Content is required"), nil
	}

	title := request.GetString("title", "")
	fontName := request.GetString("font_name", "맑은 고딕")
	fontSize := request.GetInt("font_size", 11)
	preserveFormatting := request.GetBool("preserve_formatting", true)

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil {
			hwpController = NewHwpController()
		}

		// Create new document
		err := hwpController.CreateNewDocument()
		if err != nil {
			hwpController = nil
			result = createTextResult(fmt.Sprintf("Error creating document: %v", err))
			return
		}

		// Set default font
		err = hwpController.SetFontStyle(fontName, fontSize, false, false, false)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error setting font: %v", err))
			return
		}

		// Insert title if provided
		if title != "" {
			err = hwpController.SetFontStyle(fontName, fontSize+4, true, false, false)
			if err != nil {
				result = createTextResult(fmt.Sprintf("Error setting title font: %v", err))
				return
			}
			
			err = hwpController.InsertText(title, false)
			if err != nil {
				result = createTextResult(fmt.Sprintf("Error inserting title: %v", err))
				return
			}
			
			err = hwpController.InsertParagraph()
			if err != nil {
				result = createTextResult(fmt.Sprintf("Error inserting paragraph: %v", err))
				return
			}
			
			err = hwpController.InsertParagraph()
			if err != nil {
				result = createTextResult(fmt.Sprintf("Error inserting paragraph: %v", err))
				return
			}

			// Reset font to normal
			err = hwpController.SetFontStyle(fontName, fontSize, false, false, false)
			if err != nil {
				result = createTextResult(fmt.Sprintf("Error resetting font: %v", err))
				return
			}
		}

		// Insert content
		err = hwpController.InsertText(content, preserveFormatting)
		if err != nil {
			result = createTextResult(fmt.Sprintf("Error inserting content: %v", err))
			return
		}

		result = createTextResult("Document created successfully from text")
	})
	
	return result, nil
}

func handleHwpCreateCompleteDocument(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	specStr := request.GetString("spec", "")
	if specStr == "" {
		return createTextResult("Error: Document specification is required"), nil
	}

	var result *mcp.CallToolResult
	
	executeHWPOperation(func() {
		if hwpController == nil {
			hwpController = NewHwpController()
		}

		var spec map[string]interface{}
		if err := json.Unmarshal([]byte(specStr), &spec); err != nil {
			result = createTextResult(fmt.Sprintf("Error: Failed to parse spec JSON - %v", err))
			return
		}

		// Create new document
		err := hwpController.CreateNewDocument()
		if err != nil {
			hwpController = nil
			result = createTextResult(fmt.Sprintf("Error creating document: %v", err))
			return
		}

		docType, _ := spec["type"].(string)
		
		switch docType {
		case "report":
			err = createReportDocument(spec)
		case "letter":
			err = createLetterDocument(spec)
		case "memo":
			err = createMemoDocument(spec)
		default:
			err = createGenericDocument(spec)
		}

		if err != nil {
			result = createTextResult(fmt.Sprintf("Error creating %s document: %v", docType, err))
			return
		}

		result = createTextResult(fmt.Sprintf("Complete %s document created successfully", docType))
	})
	
	return result, nil
}

func createReportDocument(spec map[string]interface{}) error {
	title, _ := spec["title"].(string)
	author, _ := spec["author"].(string)
	date, _ := spec["date"].(string)
	sections, _ := spec["sections"].([]interface{})

	// Title
	if err := hwpController.SetFontStyle("맑은 고딕", 18, true, false, false); err != nil {
		return err
	}
	if err := hwpController.InsertText(title, false); err != nil {
		return err
	}
	if err := hwpController.InsertParagraph(); err != nil {
		return err
	}

	// Author and date
	if err := hwpController.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if author != "" {
		if err := hwpController.InsertText(fmt.Sprintf("작성자: %s", author), false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}
	if date != "" {
		if err := hwpController.InsertText(fmt.Sprintf("작성일: %s", date), false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}
	if err := hwpController.InsertParagraph(); err != nil {
		return err
	}

	// Sections
	for _, sectionInterface := range sections {
		section, ok := sectionInterface.(map[string]interface{})
		if !ok {
			continue
		}
		
		sectionTitle, _ := section["title"].(string)
		sectionContent, _ := section["content"].(string)
		
		// Section title
		if err := hwpController.SetFontStyle("맑은 고딕", 14, true, false, false); err != nil {
			return err
		}
		if err := hwpController.InsertText(sectionTitle, false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}

		// Section content
		if err := hwpController.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
			return err
		}
		if err := hwpController.InsertText(sectionContent, true); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}

	return nil
}

func createLetterDocument(spec map[string]interface{}) error {
	recipient, _ := spec["recipient"].(string)
	sender, _ := spec["sender"].(string)
	date, _ := spec["date"].(string)
	subject, _ := spec["subject"].(string)
	body, _ := spec["body"].(string)
	closing, _ := spec["closing"].(string)

	// Date
	if err := hwpController.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if date != "" {
		if err := hwpController.InsertText(date, false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}

	// Recipient
	if recipient != "" {
		if err := hwpController.InsertText(fmt.Sprintf("%s 귀하", recipient), false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}

	// Subject
	if subject != "" {
		if err := hwpController.SetFontStyle("맑은 고딕", 12, true, false, false); err != nil {
			return err
		}
		if err := hwpController.InsertText(fmt.Sprintf("제목: %s", subject), false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}

	// Body
	if err := hwpController.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if err := hwpController.InsertText(body, true); err != nil {
		return err
	}
	if err := hwpController.InsertParagraph(); err != nil {
		return err
	}
	if err := hwpController.InsertParagraph(); err != nil {
		return err
	}

	// Closing and sender
	if closing != "" {
		if err := hwpController.InsertText(closing, false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}
	if sender != "" {
		if err := hwpController.InsertText(sender, false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}

	return nil
}

func createMemoDocument(spec map[string]interface{}) error {
	to, _ := spec["to"].(string)
	from, _ := spec["from"].(string)
	date, _ := spec["date"].(string)
	subject, _ := spec["subject"].(string)
	body, _ := spec["body"].(string)

	// Header
	if err := hwpController.SetFontStyle("맑은 고딕", 16, true, false, false); err != nil {
		return err
	}
	if err := hwpController.InsertText("메모", false); err != nil {
		return err
	}
	if err := hwpController.InsertParagraph(); err != nil {
		return err
	}
	if err := hwpController.InsertParagraph(); err != nil {
		return err
	}

	// Memo details
	if err := hwpController.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if to != "" {
		if err := hwpController.InsertText(fmt.Sprintf("받는 사람: %s", to), false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}
	if from != "" {
		if err := hwpController.InsertText(fmt.Sprintf("보내는 사람: %s", from), false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}
	if date != "" {
		if err := hwpController.InsertText(fmt.Sprintf("날짜: %s", date), false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}
	if subject != "" {
		if err := hwpController.InsertText(fmt.Sprintf("제목: %s", subject), false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}
	if err := hwpController.InsertParagraph(); err != nil {
		return err
	}

	// Body
	if err := hwpController.InsertText(body, true); err != nil {
		return err
	}

	return nil
}

func createGenericDocument(spec map[string]interface{}) error {
	title, _ := spec["title"].(string)
	content, _ := spec["content"].(string)

	// Title
	if title != "" {
		if err := hwpController.SetFontStyle("맑은 고딕", 16, true, false, false); err != nil {
			return err
		}
		if err := hwpController.InsertText(title, false); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
		if err := hwpController.InsertParagraph(); err != nil {
			return err
		}
	}

	// Content
	if err := hwpController.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if err := hwpController.InsertText(content, true); err != nil {
		return err
	}

	return nil
}

func newMCPServer() *server.MCPServer {
	mcpServer := server.NewMCPServer(
		"hwp-mcp-go",
		"1.0.0",
		server.WithToolCapabilities(true),
	)

	// Add all HWP tools
	mcpServer.AddTool(mcp.NewTool(HWP_CREATE,
		mcp.WithDescription("Create a new HWP document"),
	), handleHwpCreate)

	mcpServer.AddTool(mcp.NewTool(HWP_OPEN,
		mcp.WithDescription("Open an existing HWP document"),
		mcp.WithString("path",
			mcp.Description("File path to open"),
			mcp.Required(),
		),
	), handleHwpOpen)

	mcpServer.AddTool(mcp.NewTool(HWP_SAVE,
		mcp.WithDescription("Save the current HWP document"),
		mcp.WithString("path",
			mcp.Description("File path to save (optional)"),
		),
	), handleHwpSave)

	mcpServer.AddTool(mcp.NewTool(HWP_INSERT_TEXT,
		mcp.WithDescription("Insert text at the current cursor position"),
		mcp.WithString("text",
			mcp.Description("Text to insert"),
			mcp.Required(),
		),
		mcp.WithBoolean("preserve_linebreaks",
			mcp.Description("Preserve line breaks in text"),
		),
	), handleHwpInsertText)

	mcpServer.AddTool(mcp.NewTool(HWP_SET_FONT,
		mcp.WithDescription("Set font properties"),
		mcp.WithString("name",
			mcp.Description("Font name"),
		),
		mcp.WithNumber("size",
			mcp.Description("Font size"),
		),
		mcp.WithBoolean("bold",
			mcp.Description("Bold font"),
		),
		mcp.WithBoolean("italic",
			mcp.Description("Italic font"),
		),
		mcp.WithBoolean("underline",
			mcp.Description("Underline font"),
		),
	), handleHwpSetFont)

	mcpServer.AddTool(mcp.NewTool(HWP_INSERT_TABLE,
		mcp.WithDescription("Insert a table at the current cursor position"),
		mcp.WithNumber("rows",
			mcp.Description("Number of rows"),
			mcp.Required(),
		),
		mcp.WithNumber("cols",
			mcp.Description("Number of columns"),
			mcp.Required(),
		),
	), handleHwpInsertTable)

	mcpServer.AddTool(mcp.NewTool(HWP_INSERT_PARAGRAPH,
		mcp.WithDescription("Insert a new paragraph"),
	), handleHwpInsertParagraph)

	mcpServer.AddTool(mcp.NewTool(HWP_GET_TEXT,
		mcp.WithDescription("Get the text content of the current document"),
	), handleHwpGetText)

	mcpServer.AddTool(mcp.NewTool(HWP_CLOSE,
		mcp.WithDescription("Close the HWP document and connection"),
	), handleHwpClose)

	mcpServer.AddTool(mcp.NewTool(HWP_PING_PONG,
		mcp.WithDescription("Ping pong test function"),
		mcp.WithString("message",
			mcp.Description("Test message"),
		),
	), handleHwpPingPong)

	mcpServer.AddTool(mcp.NewTool(HWP_FILL_TABLE_WITH_DATA,
		mcp.WithDescription("Fill existing table with data"),
		mcp.WithString("data",
			mcp.Description("JSON string of 2D array data to fill"),
			mcp.Required(),
		),
		mcp.WithNumber("start_row",
			mcp.Description("Starting row number (1-based)"),
		),
		mcp.WithNumber("start_col",
			mcp.Description("Starting column number (1-based)"),
		),
		mcp.WithBoolean("has_header",
			mcp.Description("Whether first row is header"),
		),
	), handleHwpFillTableWithData)

	mcpServer.AddTool(mcp.NewTool(HWP_FILL_COLUMN_NUMBERS,
		mcp.WithDescription("Fill table column with sequential numbers"),
		mcp.WithNumber("start",
			mcp.Description("Starting number"),
		),
		mcp.WithNumber("end",
			mcp.Description("Ending number"),
		),
		mcp.WithNumber("column",
			mcp.Description("Column number to fill"),
		),
	), handleHwpFillColumnNumbers)

	mcpServer.AddTool(mcp.NewTool(HWP_CREATE_TABLE_WITH_DATA,
		mcp.WithDescription("Create a table and fill it with data"),
		mcp.WithNumber("rows",
			mcp.Description("Number of rows"),
			mcp.Required(),
		),
		mcp.WithNumber("cols",
			mcp.Description("Number of columns"),
			mcp.Required(),
		),
		mcp.WithString("data",
			mcp.Description("JSON string of 2D array data to fill (optional)"),
		),
		mcp.WithBoolean("has_header",
			mcp.Description("Whether first row is header"),
		),
	), handleHwpCreateTableWithData)

	mcpServer.AddTool(mcp.NewTool(HWP_BATCH_OPERATIONS,
		mcp.WithDescription("Execute multiple HWP operations in sequence"),
		mcp.WithString("operations",
			mcp.Description("JSON array of operations to execute"),
			mcp.Required(),
		),
	), handleHwpBatchOperations)

	mcpServer.AddTool(mcp.NewTool(HWP_CREATE_DOCUMENT_FROM_TEXT,
		mcp.WithDescription("Create a new document from text content"),
		mcp.WithString("content",
			mcp.Description("Text content for the document"),
			mcp.Required(),
		),
		mcp.WithString("title",
			mcp.Description("Document title (optional)"),
		),
		mcp.WithString("font_name",
			mcp.Description("Font name (default: 맑은 고딕)"),
		),
		mcp.WithNumber("font_size",
			mcp.Description("Font size (default: 11)"),
		),
		mcp.WithBoolean("preserve_formatting",
			mcp.Description("Preserve line breaks and formatting"),
		),
	), handleHwpCreateDocumentFromText)

	mcpServer.AddTool(mcp.NewTool(HWP_CREATE_COMPLETE_DOCUMENT,
		mcp.WithDescription("Create a complete document from specification (report, letter, memo)"),
		mcp.WithString("spec",
			mcp.Description("JSON specification for document creation"),
			mcp.Required(),
		),
	), handleHwpCreateCompleteDocument)

	return mcpServer
}

func main() {
	defer func() {
		if hwpController != nil {
			executeHWPOperation(func() {
				hwpController.Disconnect()
			})
		}
	}()

	mcpServer := newMCPServer()

	fmt.Fprintf(os.Stderr, "Starting HWP MCP Go server\n")
	
	if err := server.ServeStdio(mcpServer); err != nil {
		log.Fatalf("Server error: %v", err)
	}
}