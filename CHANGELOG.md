# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/ko/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

### Changed

### Fixed

### Removed
