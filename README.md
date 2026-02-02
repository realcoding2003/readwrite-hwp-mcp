# HWP MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io/)

한글(HWP) 문서를 AI 어시스턴트가 제어할 수 있게 해주는 MCP(Model Context Protocol) 서버입니다.

Windows 환경에서 한글 프로그램의 COM 인터페이스를 활용하여 HWP/HWPX 문서를 완벽하게 읽고, 쓰고, 편집할 수 있습니다.

## Features

- **문서 관리**: 생성, 열기, 저장, 닫기
- **텍스트 편집**: 삽입, 검색, 치환, 삭제
- **표 조작**: 생성, 셀 편집, 행/열 추가/삭제
- **서식 적용**: 글꼴, 정렬, 스타일
- **이미지 삽입**: 이미지 추가 및 크기 조정
- **내보내기**: PDF, 텍스트 변환

## Requirements

| 항목 | 요구사항 |
| ---- | -------- |
| OS | Windows 10/11 |
| 한글 | 2010 이상 (HWPFrame.HwpObject 지원) |
| Python | 3.10+ |
| MCP Client | Claude Desktop, Claude Code 등 |

## Installation

```bash
# 저장소 클론
git clone https://github.com/realcoding2003/readwrite-hwp-mcp.git
cd readwrite-hwp-mcp

# 의존성 설치
pip install -r requirements.txt
```

## Configuration

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json` 파일에 추가:

```json
{
  "mcpServers": {
    "hwp": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "C:\\path\\to\\readwrite-hwp-mcp"
    }
  }
}
```

### Claude Code (CLI)

```bash
claude mcp add hwp python -m src.server --cwd /path/to/readwrite-hwp-mcp
```

## Usage

MCP 클라이언트에서 사용 가능한 도구 예시:

```text
# 새 문서 생성
hwp_create

# 문서 열기
hwp_open path="C:\Documents\report.hwp"

# 텍스트 삽입
hwp_insert_text text="안녕하세요"

# 표 생성
hwp_create_table rows=3 cols=4

# PDF 내보내기
hwp_export_pdf path="C:\Documents\report.pdf"
```

## MCP Tools

### MVP (1단계)

| 도구 | 설명 |
| ---- | ---- |
| `hwp_connect` | 한글 프로그램 연결 |
| `hwp_create` | 새 문서 생성 |
| `hwp_open` | 문서 열기 |
| `hwp_save` | 문서 저장 |
| `hwp_insert_text` | 텍스트 삽입 |
| `hwp_create_table` | 표 생성 |

전체 도구 목록은 [API 문서](docs/api/README.md)를 참조하세요.

## Documentation

| 문서 | 설명 |
| ---- | ---- |
| [기획서](docs/planning/hwp-mcp-planning.md) | 프로젝트 기획 및 로드맵 |
| [아키텍처](docs/architecture/overview.md) | 시스템 구조 |
| [배포 가이드](docs/guides/deployment.md) | 설치 및 실행 |
| [API 문서](docs/api/README.md) | MCP 도구 상세 |

## Development

```bash
# 개발 환경 설정
pip install -r requirements-dev.txt

# 테스트 실행
pytest

# 린트
ruff check src/
```

## Contributing

기여를 환영합니다! 다음 단계를 따라주세요:

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'feat: add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

커밋 메시지는 [Conventional Commits](https://www.conventionalcommits.org/) 형식을 따릅니다.

## Related Projects

이 프로젝트는 다음 프로젝트들을 참고하여 개발되었습니다:

- [jkf87/hwp-mcp](https://github.com/jkf87/hwp-mcp) - HWP MCP 서버 원본
- [skerishKang/33MCP_HWP_Limone](https://github.com/skerishKang/33MCP_HWP_Limone) - 고급 HWP MCP 서버

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Hancom](https://www.hancom.com/) - 한글 프로그램 및 오토메이션 API
- [Model Context Protocol](https://modelcontextprotocol.io/) - MCP 표준
- [FastMCP](https://github.com/jlowin/fastmcp) - Python MCP 프레임워크
