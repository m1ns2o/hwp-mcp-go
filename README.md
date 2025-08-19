# HWP-MCP-Go (한글 Model Context Protocol - Go Implementation)

[![Go](https://img.shields.io/badge/Go-1.21+-blue.svg)](https://golang.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

HWP-MCP-Go는 한글 워드 프로세서(HWP)를 Claude와 같은 AI 모델이 제어할 수 있도록 해주는 Model Context Protocol(MCP) 서버의 Go 구현체입니다. mark3labs/mcp-go 라이브러리를 사용하여 Python 버전을 Go로 포팅한 프로젝트입니다.

## 주요 기능

- **문서 생성 및 관리**: 새 문서 생성, 열기, 저장 기능
- **텍스트 편집**: 텍스트 삽입, 글꼴 설정, 단락 추가
- **테이블 작업**: 테이블 생성, 데이터 채우기, 열에 연속 숫자 입력
- **고성능**: Go의 성능과 동시성 처리 능력 활용
- **크로스 플랫폼**: Windows에서 HWP COM 인터페이스 지원

## 시스템 요구사항

- Windows 운영체제
- 한글(HWP) 프로그램 설치
- Go 1.21 이상

## 설치 방법

1. 저장소 클론:
```bash
git clone <repository-url>
cd hwp-mcp-go
```

2. 의존성 설치:
```bash
go mod tidy
```

3. 빌드:
```bash
go build -o hwp-mcp-go.exe main.go
```

## 사용 방법

### Claude와 함께 사용하기

Claude 데스크톱 설정 파일에 다음과 같이 HWP-MCP-Go 서버를 등록하세요:

```json
{
  "mcpServers": {
    "hwp-go": {
      "command": "경로/hwp-mcp-go.exe"
    }
  }
}
```

### 지원되는 도구들

#### 문서 관리
- `hwp_create`: 새 문서 생성
- `hwp_open`: 문서 열기
- `hwp_save`: 문서 저장
- `hwp_close`: 문서 닫기

#### 텍스트 편집
- `hwp_insert_text`: 텍스트 삽입
- `hwp_set_font`: 글꼴 설정
- `hwp_insert_paragraph`: 단락 삽입
- `hwp_get_text`: 문서 텍스트 가져오기

#### 테이블 작업
- `hwp_insert_table`: 테이블 생성
- `hwp_fill_table_with_data`: 테이블에 데이터 채우기
- `hwp_fill_column_numbers`: 열에 연속 숫자 채우기

#### 기타
- `hwp_ping_pong`: 연결 테스트

## API 예시

### 새 문서 생성 및 텍스트 삽입
```bash
# 새 문서 생성
curl -X POST http://localhost:8080/tools/hwp_create

# 텍스트 삽입
curl -X POST http://localhost:8080/tools/hwp_insert_text \
  -d '{"text": "안녕하세요, 한글 문서입니다."}'

# 글꼴 설정
curl -X POST http://localhost:8080/tools/hwp_set_font \
  -d '{"name": "맑은 고딕", "size": 14, "bold": true}'
```

### 테이블 생성 및 데이터 입력
```bash
# 3x3 테이블 생성
curl -X POST http://localhost:8080/tools/hwp_insert_table \
  -d '{"rows": 3, "cols": 3}'

# 테이블에 데이터 채우기
curl -X POST http://localhost:8080/tools/hwp_fill_table_with_data \
  -d '{
    "data": [
      ["월", "화", "수"], 
      ["1", "2", "3"], 
      ["4", "5", "6"]
    ],
    "has_header": true
  }'

# 첫 번째 열에 1-10 숫자 채우기
curl -X POST http://localhost:8080/tools/hwp_fill_column_numbers \
  -d '{"start": 1, "end": 10, "column": 1}'
```

## 프로젝트 구조

```
hwp-mcp-go/
├── main.go              # 메인 서버 구현
├── go.mod               # Go 모듈 정의
├── README.md            # 프로젝트 문서
└── LICENSE              # 라이선스 파일
```

## 기술 스택

- **Go**: 메인 프로그래밍 언어
- **mark3labs/mcp-go**: Model Context Protocol Go 구현체
- **go-ole**: Windows COM 인터페이스 지원
- **HWP COM API**: 한글 프로그램 제어

## Python 버전과의 차이점

- **성능**: Go의 컴파일된 바이너리로 더 빠른 실행 속도
- **메모리 효율성**: Go의 가비지 컬렉터와 효율적인 메모리 관리
- **동시성**: Go의 고루틴을 활용한 비동기 처리 지원
- **배포**: 단일 실행 파일로 배포 가능 (Python 인터프리터 불필요)
- **타입 안정성**: 정적 타입 시스템으로 런타임 오류 감소

## 개발 및 기여

### 로컬 개발
```bash
# 개발 모드로 실행
go run main.go

# 테스트 실행
go test ./...

# 린터 실행 (golangci-lint 필요)
golangci-lint run
```

### 기여 방법
1. 이슈 제보 또는 기능 제안: GitHub 이슈를 사용하세요.
2. 코드 기여: Pull Request를 제출하세요.
3. 코드 스타일: Go의 표준 포맷팅을 따라주세요 (`go fmt`).

## 트러블슈팅

### HWP 연결 실패
- 한글 프로그램이 설치되어 있는지 확인
- Windows의 COM 보안 설정 확인
- 관리자 권한으로 실행 시도

### COM 인터페이스 오류
- `go-ole` 패키지가 올바르게 설치되었는지 확인
- Windows 환경에서만 실행 가능

## 라이선스

이 프로젝트는 MIT 라이선스에 따라 배포됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 관련 프로젝트

- [원본 Python 구현체](https://github.com/jkf87/hwp-mcp)
- [mark3labs/mcp-go](https://github.com/mark3labs/mcp-go)
- [HWP SDK](https://www.hancom.com/product/sdk)

## 테스트 방법

### 자동 테스트 실행

#### 1. 배치 스크립트 사용 (Windows)
```bash
test.bat
```

#### 2. Python 테스트 클라이언트
```bash
python test_mcp_client.py
```

#### 3. Go 테스트 클라이언트
```bash
# 빌드 후 실행
go build -o test-client.exe test_mcp_client.go
test-client.exe
```

### 수동 테스트

#### 1. 서버 시작
```bash
./hwp-mcp-go.exe
```

#### 2. 다른 터미널에서 JSON-RPC 요청 전송
```bash
# Initialize
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{"roots":{"listChanged":true}},"clientInfo":{"name":"test-client","version":"1.0.0"}}}' | ./hwp-mcp-go.exe

# List tools
echo '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' | ./hwp-mcp-go.exe

# Ping test
echo '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"hwp_ping_pong","arguments":{"message":"핑"}}}' | ./hwp-mcp-go.exe
```

### 테스트 항목

- ✅ MCP 프로토콜 초기화
- ✅ 도구 목록 조회
- ✅ 핑퐁 테스트 (서버 연결 확인)
- ✅ HWP 문서 생성 (한글 프로그램 필요)
- ✅ 텍스트 삽입 테스트
- ✅ HWP 연결 종료

### 문제 해결

#### COM 초기화 오류
- Windows에서만 실행 가능
- 관리자 권한으로 실행 시도
- 한글 프로그램이 설치되어 있는지 확인

#### 연결 실패
- 방화벽 설정 확인
- 한글 프로그램 버전 호환성 확인

## 변경 로그

### v1.0.0
- Python hwp-mcp의 Go 포팅 완료
- mark3labs/mcp-go 라이브러리 사용
- 모든 주요 기능 구현 (문서 관리, 텍스트 편집, 테이블 작업)
- Windows COM 인터페이스 지원
- 자동 테스트 클라이언트 추가 (Go, Python, Batch)