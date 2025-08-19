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