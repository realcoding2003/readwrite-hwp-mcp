# 배포 가이드

## 요구사항

- Node.js (권장 버전: 18 이상)
- npm 또는 yarn

## 설치

```bash
# 저장소 클론
git clone https://github.com/realcoding2003/readwrite-hwp-mcp.git
cd readwrite-hwp-mcp

# 의존성 설치
npm install
```

## 실행

```bash
# 개발 모드
npm run dev

# 프로덕션
npm run build
npm start
```

## MCP 클라이언트 설정

Claude Desktop 또는 다른 MCP 클라이언트에서 이 서버를 사용하려면:

```json
{
  "mcpServers": {
    "readwrite-hwp": {
      "command": "node",
      "args": ["path/to/readwrite-hwp-mcp/dist/index.js"]
    }
  }
}
```

## 문제 해결

(프로젝트 개발 진행에 따라 작성)
