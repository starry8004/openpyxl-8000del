# openpyxl-8000del
Excel Filter Script with openpyxl 8000del
Excel Filter Script with openpyxl 스크립트에 대한 설명.

스크립트의 주요 기능:

엑셀(.xlsx) 파일에서 검색량이 8,000 이상인 데이터만 필터링
진행 상황을 시각적으로 표시
결과를 새로운 엑셀 파일로 저장


사용된 라이브러리:
pythonCopyimport tkinter as tk  # GUI 구현
from tkinter import filedialog, messagebox, ttk  # 파일 선택, 메시지 표시, 진행바
import openpyxl  # 엑셀 파일 처리
import os  # 파일 경로 관리
from datetime import datetime  # 날짜/시간 처리

처리 과정:

GUI 창 생성
파일 선택 다이얼로그 표시
선택된 엑셀 파일 읽기
'검색량' 열 찾기
새 워크북 생성 및 헤더 복사
데이터 필터링 (검색량 8,000 이상)
필터링된 데이터 저장


진행 상황 표시:

상태 메시지로 현재 작업 표시
진행률 바로 진행도 시각화
처리 중인 행 번호 표시


파일 저장 형식:

저장 위치: 원본 파일과 동일한 폴더
파일명 형식: "원본파일명_8000del_YYYY-MM-DD_HH-MM.xlsx"
예: "상품데이터_8000del_2024-12-16_14-30.xlsx"


결과 정보 표시:

전체 데이터 수
필터링된 데이터 수
저장된 파일 경로


오류 처리:

파일 읽기/쓰기 오류 처리
검색량 데이터 변환 오류 처리
'검색량' 열 미존재 시 오류 처리


사용 방법:

스크립트 실행
처리할 엑셀 파일 선택
진행 상황 확인
완료 메시지 확인


필요한 설치:
bashCopypip install openpyxl


이 스크립트는 대용량 엑셀 파일도 처리할 수 있으며, 진행 상황을 실시간으로 확인할 수 있어 사용자 친화적입니다.
