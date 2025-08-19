package handlers

import (
	"context"
	"encoding/json"
	"fmt"

	"hwp-mcp-go/hwp-mcp-server/internal/hwp"

	"github.com/mark3labs/mcp-go/mcp"
)

// Tool names for advanced document creation
const (
	HWP_CREATE_COMPLETE_DOCUMENT = "hwp_create_complete_document"
)

// Advanced document creation tool handlers

func HandleHwpCreateCompleteDocument(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	specStr := request.GetString("spec", "")
	if specStr == "" {
		return hwp.CreateTextResult("Error: Document specification is required"), nil
	}

	var result *mcp.CallToolResult

	hwp.ExecuteHWPOperation(func() {
		controller := hwp.GetGlobalController()
		if controller == nil {
			controller = hwp.NewController()
			hwp.SetGlobalController(controller)
		}

		var spec map[string]interface{}
		if err := json.Unmarshal([]byte(specStr), &spec); err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error: Failed to parse spec JSON - %v", err))
			return
		}

		// Create new document
		err := controller.CreateNewDocument()
		if err != nil {
			hwp.SetGlobalController(nil)
			result = hwp.CreateTextResult(fmt.Sprintf("Error creating document: %v", err))
			return
		}

		docType, _ := spec["type"].(string)

		switch docType {
		case "report":
			err = createReportDocument(controller, spec)
		case "letter":
			err = createLetterDocument(controller, spec)
		case "memo":
			err = createMemoDocument(controller, spec)
		default:
			err = createGenericDocument(controller, spec)
		}

		if err != nil {
			result = hwp.CreateTextResult(fmt.Sprintf("Error creating %s document: %v", docType, err))
			return
		}

		result = hwp.CreateTextResult(fmt.Sprintf("Complete %s document created successfully", docType))
	})

	return result, nil
}

// Document creation helper functions

func createReportDocument(controller *hwp.Controller, spec map[string]interface{}) error {
	title, _ := spec["title"].(string)
	author, _ := spec["author"].(string)
	date, _ := spec["date"].(string)
	sections, _ := spec["sections"].([]interface{})

	// Title
	if err := controller.SetFontStyle("맑은 고딕", 18, true, false, false); err != nil {
		return err
	}
	if err := controller.InsertText(title, false); err != nil {
		return err
	}
	if err := controller.InsertParagraph(); err != nil {
		return err
	}

	// Author and date
	if err := controller.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if author != "" {
		if err := controller.InsertText(fmt.Sprintf("작성자: %s", author), false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}
	if date != "" {
		if err := controller.InsertText(fmt.Sprintf("작성일: %s", date), false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}
	if err := controller.InsertParagraph(); err != nil {
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
		if err := controller.SetFontStyle("맑은 고딕", 14, true, false, false); err != nil {
			return err
		}
		if err := controller.InsertText(sectionTitle, false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}

		// Section content
		if err := controller.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
			return err
		}
		if err := controller.InsertText(sectionContent, true); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}

	return nil
}

func createLetterDocument(controller *hwp.Controller, spec map[string]interface{}) error {
	recipient, _ := spec["recipient"].(string)
	sender, _ := spec["sender"].(string)
	date, _ := spec["date"].(string)
	subject, _ := spec["subject"].(string)
	body, _ := spec["body"].(string)
	closing, _ := spec["closing"].(string)

	// Date
	if err := controller.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if date != "" {
		if err := controller.InsertText(date, false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}

	// Recipient
	if recipient != "" {
		if err := controller.InsertText(fmt.Sprintf("%s 귀하", recipient), false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}

	// Subject
	if subject != "" {
		if err := controller.SetFontStyle("맑은 고딕", 12, true, false, false); err != nil {
			return err
		}
		if err := controller.InsertText(fmt.Sprintf("제목: %s", subject), false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}

	// Body
	if err := controller.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if err := controller.InsertText(body, true); err != nil {
		return err
	}
	if err := controller.InsertParagraph(); err != nil {
		return err
	}
	if err := controller.InsertParagraph(); err != nil {
		return err
	}

	// Closing and sender
	if closing != "" {
		if err := controller.InsertText(closing, false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}
	if sender != "" {
		if err := controller.InsertText(sender, false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}

	return nil
}

func createMemoDocument(controller *hwp.Controller, spec map[string]interface{}) error {
	to, _ := spec["to"].(string)
	from, _ := spec["from"].(string)
	date, _ := spec["date"].(string)
	subject, _ := spec["subject"].(string)
	body, _ := spec["body"].(string)

	// Header
	if err := controller.SetFontStyle("맑은 고딕", 16, true, false, false); err != nil {
		return err
	}
	if err := controller.InsertText("메모", false); err != nil {
		return err
	}
	if err := controller.InsertParagraph(); err != nil {
		return err
	}
	if err := controller.InsertParagraph(); err != nil {
		return err
	}

	// Memo details
	if err := controller.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if to != "" {
		if err := controller.InsertText(fmt.Sprintf("받는 사람: %s", to), false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}
	if from != "" {
		if err := controller.InsertText(fmt.Sprintf("보내는 사람: %s", from), false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}
	if date != "" {
		if err := controller.InsertText(fmt.Sprintf("날짜: %s", date), false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}
	if subject != "" {
		if err := controller.InsertText(fmt.Sprintf("제목: %s", subject), false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}
	if err := controller.InsertParagraph(); err != nil {
		return err
	}

	// Body
	if err := controller.InsertText(body, true); err != nil {
		return err
	}

	return nil
}

func createGenericDocument(controller *hwp.Controller, spec map[string]interface{}) error {
	title, _ := spec["title"].(string)
	content, _ := spec["content"].(string)

	// Title
	if title != "" {
		if err := controller.SetFontStyle("맑은 고딕", 16, true, false, false); err != nil {
			return err
		}
		if err := controller.InsertText(title, false); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
		if err := controller.InsertParagraph(); err != nil {
			return err
		}
	}

	// Content
	if err := controller.SetFontStyle("맑은 고딕", 11, false, false, false); err != nil {
		return err
	}
	if err := controller.InsertText(content, true); err != nil {
		return err
	}

	return nil
}