package handlers

import (
	"context"
	"encoding/json"
	"fmt"
	"strconv"

	"hwp-mcp-go/hwp-mcp-server/internal/hwp"

	"github.com/mark3labs/mcp-go/mcp"
)

// Tool names for table operations
const (
	HWP_INSERT_TABLE           = "hwp_insert_table"
	HWP_FILL_TABLE_WITH_DATA   = "hwp_fill_table_with_data"
	HWP_FILL_COLUMN_NUMBERS    = "hwp_fill_column_numbers"
	HWP_CREATE_TABLE_WITH_DATA = "hwp_create_table_with_data"
	// Table manipulation tools
	HWP_INSERT_LEFT_COLUMN     = "hwp_insert_left_column"
	HWP_INSERT_RIGHT_COLUMN    = "hwp_insert_right_column"
	HWP_INSERT_UPPER_ROW       = "hwp_insert_upper_row"
	HWP_INSERT_LOWER_ROW       = "hwp_insert_lower_row"
	HWP_MOVE_TO_LEFT_CELL      = "hwp_move_to_left_cell"
	HWP_MOVE_TO_RIGHT_CELL     = "hwp_move_to_right_cell"
	HWP_MOVE_TO_UPPER_CELL     = "hwp_move_to_upper_cell"
	HWP_MOVE_TO_LOWER_CELL     = "hwp_move_to_lower_cell"
	HWP_MERGE_TABLE_CELLS      = "hwp_merge_table_cells"
	HWP_MERGE_TABLES           = "hwp_merge_tables"
)

// Table operation tool handlers

func HandleHwpInsertTable(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	rows := request.GetInt("rows", 0)
	cols := request.GetInt("cols", 0)

	if rows <= 0 || cols <= 0 {
		return hwp.CreateTextResult("Error: Valid rows and cols are required"), nil
	}

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := controller.InsertTable(rows, cols)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult(fmt.Sprintf("Table created (%dx%d)", rows, cols))
	})

	return result, nil
}

func HandleHwpFillTableWithData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	dataStr := request.GetString("data", "")
	if dataStr == "" {
		return hwp.CreateTextResult("Error: Data is required"), nil
	}

	startRow := request.GetInt("start_row", 1)
	startCol := request.GetInt("start_col", 1)
	hasHeader := request.GetBool("has_header", false)

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Parse JSON data
		var tableData [][]string
		var jsonData [][]interface{}
		if err := json.Unmarshal([]byte(dataStr), &jsonData); err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: Failed to parse JSON data - %v", err))
			return
		}

		for _, rowInterface := range jsonData {
			var row []string
			for _, cell := range rowInterface {
				row = append(row, fmt.Sprintf("%v", cell))
			}
			tableData = append(tableData, row)
		}

		err := controller.FillTableWithData(tableData, startRow, startCol, hasHeader)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Table data filled successfully")
	})

	return result, nil
}

func HandleHwpFillColumnNumbers(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	start := request.GetInt("start", 1)
	end := request.GetInt("end", 10)
	column := request.GetInt("column", 1)

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Move to table beginning
		controller.MoveToTableCell("col_begin")

		// Move to specified column
		for i := 0; i < column-1; i++ {
			controller.MoveToTableCell("right")
		}

		// Fill numbers
		for num := start; num <= end; num++ {
			controller.SelectTableCell()
			controller.DeleteTableCellContent()

			// Insert text directly using controller's method
			controller.InsertText(strconv.Itoa(num), false)

			if num < end {
				controller.MoveToTableCell("lower")
			}
		}

		result = hwp.CreateTextResult(fmt.Sprintf("Column %d filled with numbers %d~%d", column, start, end))
	})

	return result, nil
}

func HandleHwpCreateTableWithData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	rows := request.GetInt("rows", 0)
	cols := request.GetInt("cols", 0)
	dataStr := request.GetString("data", "")
	hasHeader := request.GetBool("has_header", false)

	if rows <= 0 || cols <= 0 {
		return hwp.CreateTextResult("Error: Valid rows and cols are required"), nil
	}

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Create table first
		err := controller.InsertTable(rows, cols)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error creating table: %v", err))
			return
		}

		// Fill with data if provided
		if dataStr != "" {
			var tableData [][]string
			var jsonData [][]interface{}
			if err := json.Unmarshal([]byte(dataStr), &jsonData); err != nil {
				result = hwp.CreateTextResult(fmt.Sprintf("Error: Failed to parse JSON data - %v", err))
				return
			}

			for _, rowInterface := range jsonData {
				var row []string
				for _, cell := range rowInterface {
					row = append(row, fmt.Sprintf("%v", cell))
				}
				tableData = append(tableData, row)
			}

			err = controller.FillTableWithData(tableData, 1, 1, hasHeader)
			if err != nil {
				result = hwp.CreateTextResult(fmt.Sprintf("Error filling table: %v", err))
				return
			}
		}

		result = hwp.CreateTextResult(fmt.Sprintf("Table created (%dx%d) and filled with data", rows, cols))
	})

	return result, nil
}

// Table manipulation handlers

func HandleHwpInsertLeftColumn(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Insert left column
		err := controller.InsertTableColumn("left")
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Left column inserted successfully")
	})

	return result, nil
}

func HandleHwpInsertRightColumn(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Insert right column
		err := controller.InsertTableColumn("right")
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Right column inserted successfully")
	})

	return result, nil
}

func HandleHwpInsertUpperRow(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Insert upper row
		err := controller.InsertTableRow("upper")
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Upper row inserted successfully")
	})

	return result, nil
}

func HandleHwpInsertLowerRow(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Insert lower row
		err := controller.InsertTableRow("lower")
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Lower row inserted successfully")
	})

	return result, nil
}

func HandleHwpMoveToLeftCell(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Move to left cell
		err := controller.MoveToTableCell("left")
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Moved to left cell")
	})

	return result, nil
}

func HandleHwpMoveToRightCell(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Move to right cell
		err := controller.MoveToTableCell("right")
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Moved to right cell")
	})

	return result, nil
}

func HandleHwpMoveToUpperCell(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Move to upper cell
		err := controller.MoveToTableCell("upper")
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Moved to upper cell")
	})

	return result, nil
}

func HandleHwpMoveToLowerCell(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Move to lower cell
		err := controller.MoveToTableCell("lower")
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Moved to lower cell")
	})

	return result, nil
}

func HandleHwpMergeTableCells(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Merge selected cells
		err := controller.MergeTableCells()
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Table cells merged successfully")
	})

	return result, nil
}

func HandleHwpMergeTables(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		// Merge adjacent tables
		err := controller.MergeTables()
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Tables merged successfully")
	})

	return result, nil
}