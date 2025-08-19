package handlers

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"hwp-mcp-go/hwp-mcp-server/internal/hwp"

	"github.com/mark3labs/mcp-go/mcp"
)

// Tool names for text operations
const (
	HWP_INSERT_TEXT               = "hwp_insert_text"
	HWP_SET_FONT                  = "hwp_set_font"
	HWP_INSERT_PARAGRAPH          = "hwp_insert_paragraph"
	HWP_BATCH_OPERATIONS          = "hwp_batch_operations"
	HWP_CREATE_DOCUMENT_FROM_TEXT = "hwp_create_document_from_text"
	HWP_INSERT_IMAGE              = "hwp_insert_image"
	HWP_INSERT_PICTURE            = "hwp_insert_picture"
)

// Text manipulation tool handlers

func HandleHwpInsertText(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	text := request.GetString("text", "")
	if text == "" {
		return hwp.CreateTextResult("Error: Text is required"), nil
	}

	preserveLinebreaks := request.GetBool("preserve_linebreaks", true)

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := controller.InsertText(text, preserveLinebreaks)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Text inserted successfully")
	})

	return result, nil
}

func HandleHwpSetFont(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	name := request.GetString("name", "")
	size := request.GetInt("size", 0)
	bold := request.GetBool("bold", false)
	italic := request.GetBool("italic", false)
	underline := request.GetBool("underline", false)

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := controller.SetFontStyle(name, size, bold, italic, underline)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Font set successfully")
	})

	return result, nil
}

func HandleHwpInsertParagraph(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := controller.InsertParagraph()
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult("Paragraph inserted successfully")
	})

	return result, nil
}

func HandleHwpBatchOperations(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	operationsStr := request.GetString("operations", "")
	if operationsStr == "" {
		return hwp.CreateTextResult("Error: Operations list is required"), nil
	}

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		var operations []map[string]interface{}
		if err := json.Unmarshal([]byte(operationsStr), &operations); err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: Failed to parse operations JSON - %v", err))
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
				err = controller.InsertText(text, preserveLinebreaks)
			case "insert_paragraph":
				err = controller.InsertParagraph()
			case "set_font":
				name, _ := op["name"].(string)
				size := int(op["size"].(float64))
				bold, _ := op["bold"].(bool)
				italic, _ := op["italic"].(bool)
				underline, _ := op["underline"].(bool)
				err = controller.SetFontStyle(name, size, bold, italic, underline)
			case "insert_table":
				rows := int(op["rows"].(float64))
				cols := int(op["cols"].(float64))
				err = controller.InsertTable(rows, cols)
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
			"results":          results,
		})
		result = hwp.CreateTextResult(string(resultJSON))
	})

	return result, nil
}

func HandleHwpCreateDocumentFromText(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	content := request.GetString("content", "")
	if content == "" {
		return hwp.CreateTextResult("Error: Content is required"), nil
	}

	title := request.GetString("title", "")
	fontName := request.GetString("font_name", "맑은 고딕")
	fontSize := request.GetInt("font_size", 11)
	preserveFormatting := request.GetBool("preserve_formatting", true)

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil {
			controller = hwp.NewController()
			hwp.SetGlobalController(controller)
		}

		// Create new document
		err := controller.CreateNewDocument()
		if err != nil {
			hwp.SetGlobalController(nil)
			result = hwp.CreateTextResult(fmt.Sprintf("Error creating document: %v", err))
			return
		}

		// Set default font
		err = controller.SetFontStyle(fontName, fontSize, false, false, false)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error setting font: %v", err))
			return
		}

		// Insert title if provided
		if title != "" {
			err = controller.SetFontStyle(fontName, fontSize+4, true, false, false)
			if err != nil {
				result = hwp.CreateTextResult(fmt.Sprintf("Error setting title font: %v", err))
				return
			}

			err = controller.InsertText(title, false)
			if err != nil {
				result = hwp.CreateTextResult(fmt.Sprintf("Error inserting title: %v", err))
				return
			}

			err = controller.InsertParagraph()
			if err != nil {
				result = hwp.CreateTextResult(fmt.Sprintf("Error inserting paragraph: %v", err))
				return
			}

			err = controller.InsertParagraph()
			if err != nil {
				result = hwp.CreateTextResult(fmt.Sprintf("Error inserting paragraph: %v", err))
				return
			}

			// Reset font to normal
			err = controller.SetFontStyle(fontName, fontSize, false, false, false)
			if err != nil {
				result = hwp.CreateTextResult(fmt.Sprintf("Error resetting font: %v", err))
				return
			}
		}

		// Insert content
		err = controller.InsertText(content, preserveFormatting)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error inserting content: %v", err))
			return
		}

		result = hwp.CreateTextResult("Document created successfully from text")
	})

	return result, nil
}

func HandleHwpInsertImage(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	path := request.GetString("path", "")
	if path == "" {
		return hwp.CreateTextResult("Error: Image path is required"), nil
	}

	var width, height *int
	if w := request.GetInt("width", 0); w > 0 {
		width = &w
	}
	if h := request.GetInt("height", 0); h > 0 {
		height = &h
	}
	useOriginalSize := request.GetBool("use_original_size", true)
	embedded := request.GetBool("embedded", true)
	reverse := request.GetBool("reverse", false)
	watermark := request.GetBool("watermark", false)
	effect := request.GetInt("effect", 0)

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := controller.InsertImage(path, width, height, useOriginalSize, embedded, reverse, watermark, effect)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		// Generate size info string
		var sizeInfo string
		if useOriginalSize {
			sizeInfo = "original size"
		} else if width != nil && height != nil {
			sizeInfo = fmt.Sprintf("%dx%d", *width, *height)
		} else {
			sizeInfo = "custom size"
		}

		// Generate effect info
		effectNames := []string{"normal", "grayscale", "black&white"}
		var effectInfo string
		if effect >= 0 && effect < len(effectNames) {
			effectInfo = effectNames[effect]
		} else {
			effectInfo = "unknown"
		}

		// Generate options info
		var options []string
		if reverse {
			options = append(options, "reversed")
		}
		if watermark {
			options = append(options, "watermark")
		}
		if !embedded {
			options = append(options, "linked")
		}

		var optionsInfo string
		if len(options) > 0 {
			optionsInfo = fmt.Sprintf(", %s", strings.Join(options, ", "))
		}

		result = hwp.CreateTextResult(fmt.Sprintf("Image inserted successfully: %s (%s, %s effect%s)",
			path, sizeInfo, effectInfo, optionsInfo))
	})

	return result, nil
}

func HandleHwpInsertPicture(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	// This is a compatibility function that redirects to hwp_insert_image
	imagePath := request.GetString("image_path", "")
	if imagePath == "" {
		return hwp.CreateTextResult("Error: Image path is required"), nil
	}

	embedded := request.GetBool("embedded", true)
	sizeOption := request.GetInt("size_option", 1)
	reverse := request.GetBool("reverse", false)
	watermark := request.GetBool("watermark", false)
	effect := request.GetInt("effect", 0)
	var width, height *int
	if w := request.GetInt("width", 0); w > 0 {
		width = &w
	}
	if h := request.GetInt("height", 0); h > 0 {
		height = &h
	}

	useOriginalSize := (sizeOption == 0)

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil || !controller.IsRunning() || controller.GetHwp() == nil {
			result = hwp.CreateTextResult("Error: No HWP document is open. Please create or open a document first.")
			return
		}

		err := controller.InsertImage(imagePath, width, height, useOriginalSize, embedded, reverse, watermark, effect)
		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: %v", err))
			return
		}

		result = hwp.CreateTextResult(fmt.Sprintf("Picture inserted successfully: %s", imagePath))
	})

	return result, nil
}