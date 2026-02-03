# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/ko/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-02-03

### Added

- HWP MCP 서버 MVP 구현 (server)
  - HwpController: COM 인터페이스 래퍼 (connect, disconnect, 문서 조작)
  - 문서 관리 도구 7개 (hwp_connect, hwp_create, hwp_open, hwp_save, hwp_save_as, hwp_close, hwp_get_info)
  - 텍스트 편집 도구 6개 (hwp_insert_text, hwp_get_text, hwp_find_text, hwp_replace_text, hwp_insert_paragraph, hwp_move_cursor)
  - 표 조작 도구 5개 (hwp_create_table, hwp_get_cell, hwp_set_cell, hwp_insert_row, hwp_delete_row)
  - 서식 도구 3개 (hwp_set_font, hwp_set_align, hwp_export_pdf)
- 프로젝트 기반 구조 (pyproject.toml, requirements.txt)
- 유틸리티 모듈 (path_utils, validation)
- 테스트 프레임워크 설정 (pytest)
- 하이브리드 백엔드 아키텍처 (backends)
  - BaseBackend: 공통 인터페이스 추상 클래스
  - ComBackend: Windows COM 인터페이스 (HWP/HWPX 모두 지원)
  - HwpxBackend: 크로스 플랫폼 HWPX 직접 처리
  - 자동 백엔드 감지 및 선택 (factory)
- HWPX 파일 처리 모듈 (hwpx)
  - HwpxDocument: 메모리 내 문서 표현
  - HwpxReader: HWPX 파일 파싱 (ZIP+XML)
  - HwpxWriter: HWPX 파일 생성 (원자적 쓰기)
