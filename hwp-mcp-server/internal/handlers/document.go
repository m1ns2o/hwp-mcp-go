package handlers

import (
	"context"
	"fmt"

	"hwp-mcp-go/hwp-mcp-server/internal/hwp"

	"github.com/mark3labs/mcp-go/mcp"
)

// Tool names for document management
const (
	HWP_CREATE    = "hwp_create"
	HWP_OPEN      = "hwp_open"
	HWP_SAVE      = "hwp_save"
	HWP_CLOSE     = "hwp_close"
	HWP_GET_TEXT  = "hwp_get_text"
	HWP_PING_PONG = "hwp_ping_pong"
)

// Document management tool handlers

func HandleHwpCreate(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil {
			controller = hwp.NewController()
			hwp.SetGlobalController(controller)
		}

		err := controller.CreateNewDocument()
		if err != nil {
			// Reset controller on error
			hwp.SetGlobalController(nil)
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("New document created successfully")
	})

	return result, nil
}

func HandleHwpOpen(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	path := request.GetString("path", "")
	if path == "" {
		return hwp.CreateTextResult("Error: File path is required"), nil
	}

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil {
			controller = hwp.NewController()
			hwp.SetGlobalController(controller)
		}

		err := controller.OpenDocument(path)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult(fmt.Sprintf("Document opened: %s", path))
	})

	return result, nil
}

func HandleHwpSave(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	path := request.GetString("path", "")

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := controller.SaveDocument(path)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		if path != "" {
			result = hwp.CreateTextResult(fmt.Sprintf("Document saved to: %s", path))
		} else {
			result = hwp.CreateTextResult("Document saved successfully")
		}
	})

	return result, nil
}

func HandleHwpGetText(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		text, err := controller.GetText()
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult(text)
	})

	return result, nil
}

func HandleHwpClose(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil {
			result = hwp.CreateTextResult("HWP is already closed")
			return
		}

		err := controller.Disconnect()
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		hwp.SetGlobalController(nil)
		result = hwp.CreateTextResult("HWP connection closed successfully")
	})

	return result, nil
}

func HandleHwpPingPong(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
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

	resultJSON := fmt.Sprintf(`{"response":"%s","original_message":"%s","timestamp":"2024-12-19 15:04:05"}`,
		response, message)
	return hwp.CreateTextResult(resultJSON), nil
}