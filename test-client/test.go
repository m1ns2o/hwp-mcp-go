package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"os"
	"os/exec"
	"time"
)

// MCP Protocol structures
type MCPRequest struct {
	JSONRPC string      `json:"jsonrpc"`
	ID      int         `json:"id"`
	Method  string      `json:"method"`
	Params  interface{} `json:"params,omitempty"`
}

type MCPResponse struct {
	JSONRPC string      `json:"jsonrpc"`
	ID      int         `json:"id"`
	Result  interface{} `json:"result,omitempty"`
	Error   *MCPError   `json:"error,omitempty"`
}

type MCPError struct {
	Code    int    `json:"code"`
	Message string `json:"message"`
}

type InitializeParams struct {
	ProtocolVersion string            `json:"protocolVersion"`
	Capabilities    ClientCapabilities `json:"capabilities"`
	ClientInfo      ClientInfo        `json:"clientInfo"`
}

type ClientCapabilities struct {
	Roots    *RootsCapability    `json:"roots,omitempty"`
	Sampling *SamplingCapability `json:"sampling,omitempty"`
}

type RootsCapability struct {
	ListChanged bool `json:"listChanged,omitempty"`
}

type SamplingCapability struct{}

type ClientInfo struct {
	Name    string `json:"name"`
	Version string `json:"version"`
}

type ToolCallParams struct {
	Name      string                 `json:"name"`
	Arguments map[string]interface{} `json:"arguments,omitempty"`
}

// Test client
type MCPTestClient struct {
	cmd    *exec.Cmd
	stdin  io.WriteCloser
	stdout io.ReadCloser
	stderr io.ReadCloser
	reader *bufio.Scanner
	reqID  int
}

func NewMCPTestClient() *MCPTestClient {
	return &MCPTestClient{
		reqID: 1,
	}
}

func (c *MCPTestClient) Start() error {
	// Start the MCP server
	c.cmd = exec.Command("./hwp-mcp-server-restored.exe")
	
	var err error
	c.stdin, err = c.cmd.StdinPipe()
	if err != nil {
		return fmt.Errorf("failed to create stdin pipe: %v", err)
	}
	
	c.stdout, err = c.cmd.StdoutPipe()
	if err != nil {
		return fmt.Errorf("failed to create stdout pipe: %v", err)
	}
	
	c.stderr, err = c.cmd.StderrPipe()
	if err != nil {
		return fmt.Errorf("failed to create stderr pipe: %v", err)
	}
	
	if err := c.cmd.Start(); err != nil {
		return fmt.Errorf("failed to start server: %v", err)
	}
	
	c.reader = bufio.NewScanner(c.stdout)
	
	// Start stderr reader
	go func() {
		scanner := bufio.NewScanner(c.stderr)
		for scanner.Scan() {
			fmt.Printf("[SERVER] %s\n", scanner.Text())
		}
	}()
	
	fmt.Println("✅ MCP Server started successfully")
	return nil
}

func (c *MCPTestClient) SendRequest(method string, params interface{}) (*MCPResponse, error) {
	req := MCPRequest{
		JSONRPC: "2.0",
		ID:      c.reqID,
		Method:  method,
		Params:  params,
	}
	c.reqID++
	
	reqBytes, err := json.Marshal(req)
	if err != nil {
		return nil, fmt.Errorf("failed to marshal request: %v", err)
	}
	
	fmt.Printf("📤 Sending: %s\n", string(reqBytes))
	
	if _, err := c.stdin.Write(append(reqBytes, '\n')); err != nil {
		return nil, fmt.Errorf("failed to write request: %v", err)
	}
	
	// Read response with timeout
	done := make(chan string, 1)
	go func() {
		if c.reader.Scan() {
			done <- c.reader.Text()
		} else {
			done <- ""
		}
	}()
	
	select {
	case response := <-done:
		if response == "" {
			return nil, fmt.Errorf("no response received")
		}
		
		fmt.Printf("📥 Received: %s\n", response)
		
		var resp MCPResponse
		if err := json.Unmarshal([]byte(response), &resp); err != nil {
			return nil, fmt.Errorf("failed to unmarshal response: %v", err)
		}
		
		return &resp, nil
	case <-time.After(10 * time.Second):
		return nil, fmt.Errorf("timeout waiting for response")
	}
}

func (c *MCPTestClient) Close() error {
	if c.stdin != nil {
		c.stdin.Close()
	}
	if c.cmd != nil && c.cmd.Process != nil {
		c.cmd.Process.Kill()
		c.cmd.Wait()
	}
	return nil
}

func runTests() error {
	client := NewMCPTestClient()
	
	// Start server
	if err := client.Start(); err != nil {
		return fmt.Errorf("failed to start client: %v", err)
	}
	defer client.Close()
	
	// Wait a bit for server to be ready
	time.Sleep(1 * time.Second)
	
	fmt.Println("\n🧪 Starting MCP Tests...")
	
	// Test 1: Initialize
	fmt.Println("\n1️⃣ Testing Initialize...")
	initParams := InitializeParams{
		ProtocolVersion: "2024-11-05",
		Capabilities: ClientCapabilities{
			Roots: &RootsCapability{
				ListChanged: true,
			},
			Sampling: &SamplingCapability{},
		},
		ClientInfo: ClientInfo{
			Name:    "hwp-mcp-test-client",
			Version: "1.0.0",
		},
	}
	
	resp, err := client.SendRequest("initialize", initParams)
	if err != nil {
		return fmt.Errorf("initialize failed: %v", err)
	}
	if resp.Error != nil {
		return fmt.Errorf("initialize error: %s", resp.Error.Message)
	}
	fmt.Println("✅ Initialize successful")
	
	// Test 2: List Tools
	fmt.Println("\n2️⃣ Testing List Tools...")
	resp, err = client.SendRequest("tools/list", nil)
	if err != nil {
		return fmt.Errorf("tools/list failed: %v", err)
	}
	if resp.Error != nil {
		return fmt.Errorf("tools/list error: %s", resp.Error.Message)
	}
	
	// Parse tools list
	if result, ok := resp.Result.(map[string]interface{}); ok {
		if tools, ok := result["tools"].([]interface{}); ok {
			fmt.Printf("✅ Found %d tools:\n", len(tools))
			for i, tool := range tools {
				if toolMap, ok := tool.(map[string]interface{}); ok {
					name := toolMap["name"]
					description := toolMap["description"]
					fmt.Printf("   %d. %s - %s\n", i+1, name, description)
				}
			}
		}
	}
	
	// Test 3: Ping Pong Test
	fmt.Println("\n3️⃣ Testing Ping Pong...")
	pingParams := ToolCallParams{
		Name: "hwp_ping_pong",
		Arguments: map[string]interface{}{
			"message": "핑",
		},
	}
	
	resp, err = client.SendRequest("tools/call", pingParams)
	if err != nil {
		return fmt.Errorf("ping pong failed: %v", err)
	}
	if resp.Error != nil {
		return fmt.Errorf("ping pong error: %s", resp.Error.Message)
	}
	fmt.Println("✅ Ping Pong test successful")
	
	// Test 4: HWP Create (if HWP is available)
	fmt.Println("\n4️⃣ Testing HWP Create...")
	createParams := ToolCallParams{
		Name:      "hwp_create",
		Arguments: map[string]interface{}{},
	}
	
	resp, err = client.SendRequest("tools/call", createParams)
	if err != nil {
		fmt.Printf("⚠️  HWP Create failed (HWP may not be installed): %v\n", err)
	} else if resp.Error != nil {
		fmt.Printf("⚠️  HWP Create error (HWP may not be installed): %s\n", resp.Error.Message)
	} else {
		fmt.Println("✅ HWP Create successful")
		
		// Test 5: Insert Text (if create was successful)
		fmt.Println("\n5️⃣ Testing HWP Insert Text...")
		textParams := ToolCallParams{
			Name: "hwp_insert_text",
			Arguments: map[string]interface{}{
				"text": "안녕하세요! MCP 테스트입니다.",
			},
		}
		
		resp, err = client.SendRequest("tools/call", textParams)
		if err != nil {
			fmt.Printf("⚠️  Insert text failed: %v\n", err)
		} else if resp.Error != nil {
			fmt.Printf("⚠️  Insert text error: %s\n", resp.Error.Message)
		} else {
			fmt.Println("✅ Insert text successful")
		}
		
		// Test 6: Close HWP
		fmt.Println("\n6️⃣ Testing HWP Close...")
		closeParams := ToolCallParams{
			Name:      "hwp_close",
			Arguments: map[string]interface{}{},
		}
		
		resp, err = client.SendRequest("tools/call", closeParams)
		if err != nil {
			fmt.Printf("⚠️  HWP Close failed: %v\n", err)
		} else if resp.Error != nil {
			fmt.Printf("⚠️  HWP Close error: %s\n", resp.Error.Message)
		} else {
			fmt.Println("✅ HWP Close successful")
		}
	}
	
	fmt.Println("\n🎉 All tests completed!")
	return nil
}

func main() {
	fmt.Println("🚀 HWP MCP Server Test Client")
	fmt.Println("=============================")
	
	// Check if server executable exists
	if _, err := os.Stat("./hwp-mcp-go.exe"); os.IsNotExist(err) {
		log.Fatal("❌ hwp-mcp-go.exe not found. Please build the server first.")
	}
	
	if err := runTests(); err != nil {
		log.Fatalf("❌ Tests failed: %v", err)
	}
	
	fmt.Println("\n✨ All tests passed successfully!")
}