# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

HWP-MCP-Go is a Go implementation of a Model Context Protocol (MCP) server that enables AI models like Claude to control Korean word processor (HWP) applications. This is a Go port of the original Python hwp-mcp project, providing better performance and easier deployment through a single executable.

## Key Dependencies

- `github.com/mark3labs/mcp-go` - MCP protocol implementation for Go
- `github.com/go-ole/go-ole` - Windows COM interface support for HWP interaction
- Go 1.23.2+ required

## Build Commands

```bash
# Install dependencies
go mod tidy

# Build the main server executable
go build -o hwp-mcp-go.exe ./cmd/hwp-mcp-server

# Build test client
go build -o test-client.exe ./cmd/test-client

# Run in development mode
go run ./cmd/hwp-mcp-server
```

## Testing Commands

The project includes multiple testing approaches:

```bash
# Go test client (recommended)
go run ./cmd/test-client

# Or run built executable
./test-client.exe

# Python test client (if available)
python hwp-mcp/hwp-mcp-update/hwp_mcp_stdio_server.py
```

## Code Architecture

### Project Structure

The project follows Go project layout best practices:

```
hwp-mcp-go/
├── cmd/
│   ├── hwp-mcp-server/     # Main server application
│   │   └── main.go
│   └── test-client/        # Test client application  
│       └── main.go
├── internal/
│   ├── hwp/               # HWP COM interface package
│   │   └── controller.go  # Core HWP controller and thread management
│   └── handlers/          # MCP tool handlers
│       ├── document.go    # Document management tools
│       ├── text.go        # Text manipulation tools
│       ├── table.go       # Table operation tools
│       └── advanced.go    # Complex document creation tools
├── go.mod
└── go.sum
```

### Core Components

1. **HWP Controller** (`internal/hwp/controller.go`)
   - Manages HWP COM interface connection via `Controller` struct
   - Handles single-threaded COM operations via dedicated goroutine
   - Provides methods for document operations (create, open, save, close)
   - Text manipulation (insert, font styling, paragraphs)
   - Table operations (create, fill with data, column numbering)
   - Global controller instance managed via `GetGlobalController()` and `SetGlobalController()`

2. **COM Thread Management** (`internal/hwp/controller.go`)
   - Uses `hwpOperationCh` channel for single-threaded COM operations
   - `ExecuteHWPOperation*` functions ensure all HWP calls happen on dedicated thread
   - Critical for Windows COM stability

3. **MCP Tool Handlers** (`internal/handlers/`)
   - Document tools: `document.go` - Create, open, save, close, get text, ping-pong
   - Text tools: `text.go` - Insert text, set font, paragraphs, batch operations
   - Table tools: `table.go` - Create tables, fill data, column numbering
   - Advanced tools: `advanced.go` - Complex document creation (reports, letters, memos)
   - All handlers use `hwp.ExecuteHWPOperation` to ensure thread safety

4. **Server Applications** (`cmd/`)
   - `hwp-mcp-server/main.go`: Main MCP server with tool registration
   - `test-client/main.go`: Test client for validation
   - Server setup via `newMCPServer()` function

### Tool Categories

- **Document Management**: `hwp_create`, `hwp_open`, `hwp_save`, `hwp_close`, `hwp_get_text`
- **Text Operations**: `hwp_insert_text`, `hwp_set_font`, `hwp_insert_paragraph`
- **Image Operations**: `hwp_insert_image`, `hwp_insert_picture` (compatibility)
- **Table Operations**: `hwp_insert_table`, `hwp_fill_table_with_data`, `hwp_fill_column_numbers`, `hwp_create_table_with_data`
- **Table Manipulation**: `hwp_insert_left_column`, `hwp_insert_right_column`, `hwp_insert_upper_row`, `hwp_insert_lower_row`, `hwp_move_to_left_cell`, `hwp_move_to_right_cell`, `hwp_move_to_upper_cell`, `hwp_move_to_lower_cell`, `hwp_merge_table_cells`
- **Advanced Features**: `hwp_batch_operations`, `hwp_create_document_from_text`, `hwp_create_complete_document`
- **Utility**: `hwp_ping_pong` (connection testing)

### Thread Safety Considerations

- All HWP COM operations MUST use `executeHWPOperation` wrapper
- COM initialization is handled once per dedicated thread
- Safe COM method calling with panic recovery via `safeCallMethod` and `safeGetProperty`

## Claude Desktop Integration

To use with Claude Desktop, add to your configuration:

```json
{
  "mcpServers": {
    "hwp-go": {
      "command": "path/to/hwp-mcp-go.exe"
    }
  }
}
```

## Platform Requirements

- **Windows Only**: Uses Windows COM interface to control HWP
- Korean HWP (한글) software must be installed
- Administrative privileges may be required for COM operations

## Error Handling Patterns

- COM operations wrapped in panic recovery
- Connection state validation before operations
- Graceful error messages for missing HWP installation
- Thread-safe error propagation through channels

## Development Notes

- The `hwp-mcp/` directory contains the original Python implementation for reference
- Test client validates MCP protocol compliance and basic HWP functionality
- Server runs as stdio-based process for MCP client communication