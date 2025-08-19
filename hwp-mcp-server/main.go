package main

import (
	"fmt"
	"log"
	"os"

	"hwp-mcp-go/hwp-mcp-server/internal/handlers"
	"hwp-mcp-go/hwp-mcp-server/internal/hwp"

	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
)

// newMCPServer creates and configures the MCP server with all HWP tools
func newMCPServer() *server.MCPServer {
	mcpServer := server.NewMCPServer(
		"hwp-mcp-go",
		"1.0.0",
		server.WithToolCapabilities(true),
	)

	// Document management tools
	mcpServer.AddTool(mcp.NewTool(handlers.HWP_CREATE,
		mcp.WithDescription("Create a new HWP document"),
	), handlers.HandleHwpCreate)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_OPEN,
		mcp.WithDescription("Open an existing HWP document"),
		mcp.WithString("path",
			mcp.Description("File path to open"),
			mcp.Required(),
		),
	), handlers.HandleHwpOpen)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_SAVE,
		mcp.WithDescription("Save the current HWP document"),
		mcp.WithString("path",
			mcp.Description("File path to save (optional)"),
		),
	), handlers.HandleHwpSave)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_GET_TEXT,
		mcp.WithDescription("Get the text content of the current document"),
	), handlers.HandleHwpGetText)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_CLOSE,
		mcp.WithDescription("Close the HWP document and connection"),
	), handlers.HandleHwpClose)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_PING_PONG,
		mcp.WithDescription("Ping pong test function"),
		mcp.WithString("message",
			mcp.Description("Test message"),
		),
	), handlers.HandleHwpPingPong)

	// Text manipulation tools
	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_TEXT,
		mcp.WithDescription("Insert text at the current cursor position"),
		mcp.WithString("text",
			mcp.Description("Text to insert"),
			mcp.Required(),
		),
		mcp.WithBoolean("preserve_linebreaks",
			mcp.Description("Preserve line breaks in text"),
		),
	), handlers.HandleHwpInsertText)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_SET_FONT,
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
	), handlers.HandleHwpSetFont)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_PARAGRAPH,
		mcp.WithDescription("Insert a new paragraph"),
	), handlers.HandleHwpInsertParagraph)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_BATCH_OPERATIONS,
		mcp.WithDescription("Execute multiple HWP operations in sequence"),
		mcp.WithString("operations",
			mcp.Description("JSON array of operations to execute"),
			mcp.Required(),
		),
	), handlers.HandleHwpBatchOperations)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_CREATE_DOCUMENT_FROM_TEXT,
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
	), handlers.HandleHwpCreateDocumentFromText)

	// Image insertion tools
	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_IMAGE,
		mcp.WithDescription("Insert an image at the current cursor position"),
		mcp.WithString("path",
			mcp.Description("Image file path or URL"),
			mcp.Required(),
		),
		mcp.WithNumber("width",
			mcp.Description("Image width (hwpunit)"),
		),
		mcp.WithNumber("height",
			mcp.Description("Image height (hwpunit)"),
		),
		mcp.WithBoolean("use_original_size",
			mcp.Description("Use original image size"),
		),
		mcp.WithBoolean("embedded",
			mcp.Description("Embed image in document"),
		),
		mcp.WithBoolean("reverse",
			mcp.Description("Flip image horizontally"),
		),
		mcp.WithBoolean("watermark",
			mcp.Description("Set as watermark"),
		),
		mcp.WithNumber("effect",
			mcp.Description("Image effect (0: normal, 1: grayscale, 2: black&white)"),
		),
	), handlers.HandleHwpInsertImage)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_PICTURE,
		mcp.WithDescription("Insert a picture (compatibility function)"),
		mcp.WithString("image_path",
			mcp.Description("Image file path"),
			mcp.Required(),
		),
		mcp.WithBoolean("embedded",
			mcp.Description("Embed image in document"),
		),
		mcp.WithNumber("size_option",
			mcp.Description("Size option (0: original, 1: specified)"),
		),
		mcp.WithBoolean("reverse",
			mcp.Description("Flip horizontally"),
		),
		mcp.WithBoolean("watermark",
			mcp.Description("Set as watermark"),
		),
		mcp.WithNumber("effect",
			mcp.Description("Image effect"),
		),
		mcp.WithNumber("width",
			mcp.Description("Image width"),
		),
		mcp.WithNumber("height",
			mcp.Description("Image height"),
		),
	), handlers.HandleHwpInsertPicture)

	// Table operation tools
	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_TABLE,
		mcp.WithDescription("Insert a table at the current cursor position"),
		mcp.WithNumber("rows",
			mcp.Description("Number of rows"),
			mcp.Required(),
		),
		mcp.WithNumber("cols",
			mcp.Description("Number of columns"),
			mcp.Required(),
		),
	), handlers.HandleHwpInsertTable)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_FILL_TABLE_WITH_DATA,
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
	), handlers.HandleHwpFillTableWithData)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_FILL_COLUMN_NUMBERS,
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
	), handlers.HandleHwpFillColumnNumbers)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_CREATE_TABLE_WITH_DATA,
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
	), handlers.HandleHwpCreateTableWithData)

	// Table manipulation tools
	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_LEFT_COLUMN,
		mcp.WithDescription("Insert a column to the left of the current position"),
	), handlers.HandleHwpInsertLeftColumn)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_RIGHT_COLUMN,
		mcp.WithDescription("Insert a column to the right of the current position"),
	), handlers.HandleHwpInsertRightColumn)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_UPPER_ROW,
		mcp.WithDescription("Insert a row above the current position"),
	), handlers.HandleHwpInsertUpperRow)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_INSERT_LOWER_ROW,
		mcp.WithDescription("Insert a row below the current position"),
	), handlers.HandleHwpInsertLowerRow)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_MOVE_TO_LEFT_CELL,
		mcp.WithDescription("Move cursor to the left cell"),
	), handlers.HandleHwpMoveToLeftCell)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_MOVE_TO_RIGHT_CELL,
		mcp.WithDescription("Move cursor to the right cell"),
	), handlers.HandleHwpMoveToRightCell)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_MOVE_TO_UPPER_CELL,
		mcp.WithDescription("Move cursor to the upper cell"),
	), handlers.HandleHwpMoveToUpperCell)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_MOVE_TO_LOWER_CELL,
		mcp.WithDescription("Move cursor to the lower cell"),
	), handlers.HandleHwpMoveToLowerCell)

	mcpServer.AddTool(mcp.NewTool(handlers.HWP_MERGE_TABLE_CELLS,
		mcp.WithDescription("Merge selected table cells"),
	), handlers.HandleHwpMergeTableCells)

	// Advanced document creation tools
	mcpServer.AddTool(mcp.NewTool(handlers.HWP_CREATE_COMPLETE_DOCUMENT,
		mcp.WithDescription("Create a complete document from specification (report, letter, memo)"),
		mcp.WithString("spec",
			mcp.Description("JSON specification for document creation"),
			mcp.Required(),
		),
	), handlers.HandleHwpCreateCompleteDocument)

	return mcpServer
}

func main() {
	// Cleanup on exit
	defer func() {
		controller := hwp.GetGlobalController()
		if controller != nil {
			hwp.ExecuteHWPOperation(func() {
				controller.Disconnect()
			})
		}
	}()

	// Create and configure MCP server
	mcpServer := newMCPServer()

	fmt.Fprintf(os.Stderr, "Starting HWP MCP Go server\n")

	// Start stdio-based MCP server
	if err := server.ServeStdio(mcpServer); err != nil {
		log.Fatalf("Server error: %v", err)
	}
}