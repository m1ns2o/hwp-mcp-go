"""
한글(HWP) 문서의 표 관련 기능을 제공하는 모듈
hwp_controller.py와 함께 사용됩니다.
"""

import json
import logging
import os
import time
from typing import List, Dict, Any, Optional

# Configure logging
logger = logging.getLogger("hwp-table-tools")


class HwpTableTools:
    """한글 문서의 표 관련 기능을 제공하는 클래스"""

    def __init__(self, hwp_controller=None):
        """
        초기화 함수
        
        Args:
            hwp_controller: HwpController 인스턴스
        """
        self.hwp_controller = hwp_controller

    def set_controller(self, hwp_controller):
        """
        HwpController 인스턴스 설정
        
        Args:
            hwp_controller: HwpController 인스턴스
        """
        self.hwp_controller = hwp_controller

    def insert_table(self, rows: int, cols: int) -> str:
        """
        현재 커서 위치에 표를 삽입합니다.
        
        Args:
            rows: 표의 행 수
            cols: 표의 열 수
            
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            if self.hwp_controller.insert_table(rows, cols):
                logger.info(f"Successfully inserted {rows}x{cols} table")
                return f"Table inserted with {rows} rows and {cols} columns"
            else:
                return "Error: Failed to insert table"
        except Exception as e:
            logger.error(f"Error inserting table: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def set_cell_text(self, row: int, col: int, text: str) -> str:
        """
        표의 특정 셀에 텍스트를 입력합니다.
        
        Args:
            row: 셀의 행 번호 (1부터 시작)
            col: 셀의 열 번호 (1부터 시작)
            text: 입력할 텍스트
            
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            # fill_table_cell 메서드를 사용하여 셀에 텍스트 입력
            if self.hwp_controller.fill_table_cell(row, col, text):
                logger.info(f"셀 텍스트 설정 완료: ({row}, {col})")
                return f"셀({row}, {col})에 텍스트 입력 완료"
            else:
                return f"셀({row}, {col})에 텍스트 입력 실패"
        except Exception as e:
            logger.error(f"셀 텍스트 설정 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def merge_cells(self, start_row: int, start_col: int, end_row: int, end_col: int) -> str:
        """
        표의 특정 범위의 셀을 병합합니다.
        
        Args:
            start_row: 시작 행 번호 (1부터 시작)
            start_col: 시작 열 번호 (1부터 시작)
            end_row: 종료 행 번호 (1부터 시작)
            end_col: 종료 열 번호 (1부터 시작)
            
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            # merge_table_cells 메서드를 사용하여 셀 병합
            if self.hwp_controller.merge_table_cells(start_row, start_col, end_row, end_col):
                logger.info(f"셀 병합 완료: ({start_row},{start_col}) - ({end_row},{end_col})")
                return f"셀 병합 완료 ({start_row},{start_col}) - ({end_row},{end_col})"
            else:
                return f"셀 병합 실패"
        except Exception as e:
            logger.error(f"셀 병합 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def get_cell_text(self, row: int, col: int) -> str:
        """
        표의 특정 셀의 텍스트를 가져옵니다.
        
        Args:
            row: 셀의 행 번호 (1부터 시작)
            col: 셀의 열 번호 (1부터 시작)
            
        Returns:
            str: 셀의 텍스트 내용
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            # get_table_cell_text 메서드를 사용하여 셀 텍스트 가져오기
            text = self.hwp_controller.get_table_cell_text(row, col)
            logger.info(f"셀 텍스트 가져오기 완료: ({row}, {col})")
            return text
        except Exception as e:
            logger.error(f"셀 텍스트 가져오기 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def create_table_with_data(self, rows: int, cols: int, data: str = None, has_header: bool = False) -> str:
        """
        현재 커서 위치에 표를 생성하고 데이터를 채웁니다.
        
        Args:
            rows: 표의 행 수
            cols: 표의 열 수
            data: 표에 채울 데이터 (JSON 형식의 2차원 배열 문자열, 예: '[["항목1", "항목2"], ["값1", "값2"]]')
            has_header: 첫 번째 행을 헤더로 처리할지 여부
            
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            # 표 생성
            if not self.hwp_controller.insert_table(rows, cols):
                return "Error: Failed to create table"

            # 데이터가 제공된 경우 표 채우기
            if data:
                try:
                    # 입력 데이터 로깅
                    logger.info(f"Parsing data string: {data[:100]}...")

                    # JSON 문자열을 파이썬 객체로 변환
                    data_array = json.loads(data)

                    # 데이터 구조 유효성 검사
                    if not isinstance(data_array, list):
                        return f"표는 생성되었으나 데이터가 리스트 형식이 아닙니다. 받은 데이터 타입: {type(data_array)}"

                    if len(data_array) == 0:
                        return f"표는 생성되었으나 데이터 리스트가 비어 있습니다."

                    if not all(isinstance(row, list) for row in data_array):
                        return f"표는 생성되었으나 데이터가 2차원 배열 형식이 아닙니다."

                    # 모든 문자열로 변환 (혼합 유형 데이터 처리)
                    str_data_array = [[str(cell) for cell in row] for row in data_array]

                    logger.info(f"Converted data array: {str_data_array[:2]}...")

                    # fill_table_with_data 메서드를 사용하여 데이터 채우기
                    if self.hwp_controller.fill_table_with_data(str_data_array, 1, 1, has_header):
                        return f"표 생성 및 데이터 입력 완료 ({rows}x{cols} 크기)"
                    else:
                        return f"표는 생성되었으나 데이터 입력에 실패했습니다."

                except json.JSONDecodeError as e:
                    logger.error(f"JSON 파싱 오류: {str(e)}")
                    return f"표는 생성되었으나 JSON 데이터 파싱 오류: {str(e)}"
                except Exception as data_error:
                    logger.error(f"표 데이터 입력 중 오류: {str(data_error)}", exc_info=True)
                    return f"표는 생성되었으나 데이터 입력 중 오류 발생: {str(data_error)}"

            return f"표 생성 완료 ({rows}x{cols} 크기)"
        except Exception as e:
            logger.error(f"표 생성 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def fill_table_with_data(self, data_list: List[List[str]], start_row: int = 1, start_col: int = 1,
                             has_header: bool = False) -> str:
        """
        이미 존재하는 표에 데이터를 채웁니다.
        
        Args:
            data_list: 표에 채울 2차원 데이터 리스트
            start_row: 시작 행 번호 (1부터 시작)
            start_col: 시작 열 번호 (1부터 시작)
            has_header: 첫 번째 행을 헤더로 처리할지 여부
            
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            if not data_list:
                return "Error: Data is required"

            logger.info(f"Filling table with data: {len(data_list)} rows, starting at ({start_row}, {start_col})")

            # 데이터 형식 검사 및 변환
            processed_data = []
            for row in data_list:
                if not isinstance(row, list):
                    logger.warning(f"행이 리스트 형식이 아님: {type(row)}")
                    row = [str(row)]
                processed_row = [str(cell) if cell is not None else "" for cell in row]
                processed_data.append(processed_row)

            # fill_table_with_data 메서드를 사용하여 데이터 채우기
            success = self.hwp_controller.fill_table_with_data(processed_data, start_row, start_col, has_header)

            if success:
                logger.info("표 데이터 입력 완료")
                return "표 데이터 입력 완료"
            else:
                logger.error("hwp_controller.fill_table_with_data 호출 실패")
                return "표 데이터 입력 실패"
        except Exception as e:
            logger.error(f"표 데이터 입력 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    # === 칼럼(열) 삽입 기능 ===
    def insert_left_column(self) -> str:
        """
        현재 위치 왼쪽에 칼럼(열)을 삽입합니다.
        
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            success = self.hwp_controller.run_action("TableInsertLeftColumn")
            if success:
                logger.info("왼쪽 칼럼 삽입 완료")
                return "왼쪽 칼럼 삽입 완료"
            else:
                return "왼쪽 칼럼 삽입 실패"
        except Exception as e:
            logger.error(f"왼쪽 칼럼 삽입 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def insert_right_column(self) -> str:
        """
        현재 위치 오른쪽에 칼럼(열)을 삽입합니다.
        
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            success = self.hwp_controller.run_action("TableInsertRightColumn")
            if success:
                logger.info("오른쪽 칼럼 삽입 완료")
                return "오른쪽 칼럼 삽입 완료"
            else:
                return "오른쪽 칼럼 삽입 실패"
        except Exception as e:
            logger.error(f"오른쪽 칼럼 삽입 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    # === 행(줄) 삽입 기능 ===
    def insert_upper_row(self) -> str:
        """
        현재 위치 위쪽에 행(줄)을 삽입합니다.
        
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            success = self.hwp_controller.run_action("TableInsertUpperRow")
            if success:
                logger.info("위쪽 행 삽입 완료")
                return "위쪽 행 삽입 완료"
            else:
                return "위쪽 행 삽입 실패"
        except Exception as e:
            logger.error(f"위쪽 행 삽입 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def insert_lower_row(self) -> str:
        """
        현재 위치 아래쪽에 행(줄)을 삽입합니다.
        
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            success = self.hwp_controller.run_action("TableInsertLowerRow")
            if success:
                logger.info("아래쪽 행 삽입 완료")
                return "아래쪽 행 삽입 완료"
            else:
                return "아래쪽 행 삽입 실패"
        except Exception as e:
            logger.error(f"아래쪽 행 삽입 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    # === 셀 이동 기능 ===
    def move_to_left_cell(self) -> str:
        """
        왼쪽 셀로 이동합니다.
        
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            success = self.hwp_controller.run_action("TableLeftCell")
            if success:
                logger.info("왼쪽 셀로 이동 완료")
                return "왼쪽 셀로 이동 완료"
            else:
                return "왼쪽 셀로 이동 실패"
        except Exception as e:
            logger.error(f"왼쪽 셀로 이동 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def move_to_right_cell(self) -> str:
        """
        오른쪽 셀로 이동합니다.
        
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            success = self.hwp_controller.run_action("TableRightCell")
            if success:
                logger.info("오른쪽 셀로 이동 완료")
                return "오른쪽 셀로 이동 완료"
            else:
                return "오른쪽 셀로 이동 실패"
        except Exception as e:
            logger.error(f"오른쪽 셀로 이동 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def move_to_upper_cell(self) -> str:
        """
        위쪽 셀로 이동합니다.
        
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            success = self.hwp_controller.run_action("TableUpperCell")
            if success:
                logger.info("위쪽 셀로 이동 완료")
                return "위쪽 셀로 이동 완료"
            else:
                return "위쪽 셀로 이동 실패"
        except Exception as e:
            logger.error(f"위쪽 셀로 이동 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    def move_to_lower_cell(self) -> str:
        """
        아래쪽 셀로 이동합니다.
        
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"

            success = self.hwp_controller.run_action("TableLowerCell")
            if success:
                logger.info("아래쪽 셀로 이동 완료")
                return "아래쪽 셀로 이동 완료"
            else:
                return "아래쪽 셀로 이동 실패"
        except Exception as e:
            logger.error(f"아래쪽 셀로 이동 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"

    # === 셀 병합 기능 ===
    # def merge_table_cells(self) -> str:
    #     """
    #     선택된 셀들을 병합합니다.
    #
    #     Returns:
    #         str: 결과 메시지
    #     """
    #     try:
    #         if not self.hwp_controller:
    #             return "Error: HWP Controller is not set"
    #
    #         success = self.hwp_controller.run_action("TableMergeCell")
    #         if success:
    #             logger.info("셀 병합 완료")
    #             return "셀 병합 완료"
    #         else:
    #             return "셀 병합 실패"
    #     except Exception as e:
    #         logger.error(f"셀 병합 중 오류: {str(e)}", exc_info=True)
    #         return f"Error: {str(e)}"
    #
    # def merge_cells_range(self, start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    #     """
    #     지정된 범위의 셀들을 자동으로 선택하고 병합합니다.
    #
    #     Args:
    #         start_row: 시작 행 번호 (1부터 시작)
    #         start_col: 시작 열 번호 (1부터 시작)
    #         end_row: 종료 행 번호 (1부터 시작)
    #         end_col: 종료 열 번호 (1부터 시작)
    #
    #     Returns:
    #         str: 결과 메시지
    #     """
    #     try:
    #         if not self.hwp_controller:
    #             return "Error: HWP Controller is not set"
    #
    #         logger.info(f"셀 범위 병합 시작: ({start_row},{start_col}) - ({end_row},{end_col})")
    #
    #         # HWP의 올바른 셀 범위 선택 방법 사용
    #         hwp = self.hwp_controller.hwp
    #
    #         # 1. 표 첫 번째 셀로 이동
    #         hwp.Run("TableSelTable")  # 표 전체 선택
    #         hwp.Run("Cancel")  # 선택 취소 (첫 번째 셀에 위치)
    #
    #         # 2. 시작 셀로 이동
    #         for _ in range(start_row - 1):
    #             hwp.Run("TableLowerCell")
    #         for _ in range(start_col - 1):
    #             hwp.Run("TableRightCell")
    #
    #         # 3. 시작 셀 선택
    #         hwp.Run("TableSelCell")
    #
    #         # 4. HWP의 셀 범위 선택 방법 - TableSelCellRange 시도
    #         try:
    #             # TableSelCellRange 액션이 있다면 사용
    #             result = hwp.Run("TableSelCellRange")
    #             if result is not False:
    #                 logger.info("TableSelCellRange 사용")
    #         except:
    #             logger.info("TableSelCellRange 없음 - 대안 방법 사용")
    #
    #         # 5. 범위 확장 - 정확한 방법으로 수정
    #         # 먼저 행 확장 (아래쪽으로)
    #         for i in range(end_row - start_row):
    #             try:
    #                 hwp.Run("TableSelCellExt")  # 선택 확장 모드
    #                 hwp.Run("TableLowerCell")  # 아래로 이동하면서 확장
    #             except:
    #                 # 대안: Shift를 누른 상태로 이동
    #                 hwp.Run("TableLowerCell")
    #
    #         # 그다음 열 확장 (오른쪽으로)
    #         for i in range(end_col - start_col):
    #             try:
    #                 hwp.Run("TableSelCellExt")  # 선택 확장 모드
    #                 hwp.Run("TableRightCell")  # 오른쪽으로 이동하면서 확장
    #             except:
    #                 # 대안: Shift를 누른 상태로 이동
    #                 hwp.Run("TableRightCell")
    #
    #         # 6. 병합 실행 전에 잠시 대기 (선택 상태 안정화)
    #         import time
    #         time.sleep(0.1)
    #
    #         # 7. 병합 실행
    #         success = self.hwp_controller.run_action("TableMergeCell")
    #
    #         if success:
    #             logger.info(f"셀 범위 병합 완료: ({start_row},{start_col}) - ({end_row},{end_col})")
    #             return f"셀 범위 병합 완료: ({start_row},{start_col}) - ({end_row},{end_col})"
    #         else:
    #             # 병합 실패 시 다른 방법 시도
    #             logger.warning("첫 번째 방법 실패 - 대안 방법 시도")
    #
    #             # 대안 방법: 직접 셀 좌표를 이용한 선택
    #             try:
    #                 # 선택 취소하고 다시 시작
    #                 hwp.Run("Cancel")
    #
    #                 # 시작 셀로 다시 이동
    #                 hwp.Run("TableSelTable")
    #                 hwp.Run("Cancel")
    #                 for _ in range(start_row - 1):
    #                     hwp.Run("TableLowerCell")
    #                 for _ in range(start_col - 1):
    #                     hwp.Run("TableRightCell")
    #
    #                 # 단순히 인접한 셀들만 병합 (작은 범위부터)
    #                 if start_row == end_row and start_col < end_col:
    #                     # 같은 행의 셀들 병합 (가로 병합)
    #                     hwp.Run("TableSelCell")
    #                     for _ in range(end_col - start_col):
    #                         hwp.Run("TableRightCell")
    #                     success = self.hwp_controller.run_action("TableMergeCell")
    #                 elif start_col == end_col and start_row < end_row:
    #                     # 같은 열의 셀들 병합 (세로 병합)
    #                     hwp.Run("TableSelCell")
    #                     for _ in range(end_row - start_row):
    #                         hwp.Run("TableLowerCell")
    #                     success = self.hwp_controller.run_action("TableMergeCell")
    #
    #                 if success:
    #                     return f"셀 범위 병합 완료 (대안방법): ({start_row},{start_col}) - ({end_row},{end_col})"
    #
    #             except Exception as e2:
    #                 logger.error(f"대안 방법도 실패: {str(e2)}")
    #
    #             return f"셀 범위 병합 실패: 선택 또는 병합 불가"
    #
    #     except Exception as e:
    #         logger.error(f"셀 범위 병합 중 오류: {str(e)}", exc_info=True)
    #         return f"Error: {str(e)}"
    #
    # # === 표 병합 기능 ===
    def merge_tables(self) -> str:
        """
        표를 병합합니다.
    
        Returns:
            str: 결과 메시지
        """
        try:
            if not self.hwp_controller:
                return "Error: HWP Controller is not set"
    
            success = self.hwp_controller.run_action("TableMergeTable")
            if success:
                logger.info("표 병합 완료")
                return "표 병합 완료"
            else:
                return "표 병합 실패"
        except Exception as e:
            logger.error(f"표 병합 중 오류: {str(e)}", exc_info=True)
            return f"Error: {str(e)}"


# 유틸리티 함수 - 문자열 데이터를 2차원 배열로 변환
def parse_table_data(data_str: str) -> List[List[str]]:
    """
    문자열 형태의 표 데이터를 2차원 리스트로 변환합니다.
    
    Args:
        data_str: JSON 형식의 2차원 배열 문자열
        
    Returns:
        List[List[str]]: 2차원 데이터 리스트
    """
    try:
        data = json.loads(data_str)

        # 데이터 구조 유효성 검사
        if not isinstance(data, list):
            logger.error(f"데이터가 리스트 형식이 아님: {type(data)}")
            return []

        # 모든 행이 리스트인지 확인하고 문자열로 변환
        result = []
        for row in data:
            if isinstance(row, list):
                result.append([str(cell) if cell is not None else "" for cell in row])
            else:
                # 리스트가 아닌 행은 단일 항목 리스트로 처리
                result.append([str(row)])

        return result
    except json.JSONDecodeError as e:
        logger.error(f"표 데이터 파싱 오류: {str(e)}")
        return []
