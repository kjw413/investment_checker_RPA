"""
투자실적 검증 자동화 프로그램

실제 투자이력과 전산(ERP) 상 등록된 투자이력을 비교하여
미등록 및 계정코드 누락 사항을 검출합니다.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

# tkinter는 헤드리스 환경에서 사용 불가능할 수 있음
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

# ============================================================================
# 설정 및 상수
# ============================================================================

@dataclass
class Config:
    """프로그램 설정 클래스"""
    base_dir: Path = Path(__file__).resolve().parents[1]
    erp_dir: Path | None = None
    output_dir: Path | None = None
    
    # 헤더 탐지 설정
    max_scan_rows: int = 200
    min_header_hits: int = 3
    
    # 정규화 설정
    null_tokens: frozenset[str] = frozenset({"nan", "none", "null"})
    
    # 컬럼 별칭 (헤더 탐지용)
    column_aliases: dict[str, list[str]] = None
    
    def __post_init__(self):
        # 기본 ERP 디렉터리 자동 감지: 여러 후보 폴더명을 확인합니다.
        if self.erp_dir is None:
            candidates = [
                self.base_dir / "종합정보 투자실적",
                self.base_dir / "투자리스트_종합정보시스템",
                self.base_dir / "투자리스트_실제",
            ]
            for c in candidates:
                if c.exists() and c.is_dir():
                    self.erp_dir = c
                    break
            else:
                # 후보 미발견 시 기존 경로를 기본값으로 설정
                self.erp_dir = self.base_dir / "종합정보 투자실적"

        if self.output_dir is None:
            self.output_dir = self.base_dir / "result"

        if self.column_aliases is None:
            self.column_aliases = {
                "공장": ["공장", "공장명", "사업장", "사업장명"],
                "부서": ["부서", "부서명", "사용부서", "요청부서"],
                "투자명": ["투자명", "건명", "품의명", "과제명", "내역"],
                "투자코드": ["투자코드", "코드", "erp코드", "모코드", "자코드"],
                "자본적": ["자본적", "자본", "capex"],
                "수익적": ["수익적", "수익", "opex"],
            }


# ============================================================================
# 로깅 설정
# ============================================================================

def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """로깅 설정"""
    logger = logging.getLogger("InvestmentValidator")
    logger.setLevel(level)
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger

logger = setup_logging()


# ============================================================================
# 데이터 정규화 유틸리티
# ============================================================================

class DataNormalizer:
    """데이터 정규화 담당 클래스"""
    
    def __init__(self, config: Config):
        self.config = config
    
    def normalize_code_series(self, series: pd.Series) -> pd.Series:
        """코드 컬럼 정규화 (공백 제거, null 처리)"""
        normalized = series.fillna("").astype(str)
        normalized = normalized.str.replace(r"\s+", "", regex=True)
        
        # null 토큰 처리
        lower = normalized.str.lower()
        normalized = normalized.mask(
            lower.isin(self.config.null_tokens), 
            ""
        )
        
        return normalized
    
    def to_number_series(self, series: pd.Series) -> pd.Series:
        """숫자 컬럼 변환 (콤마 제거, 숫자 변환)"""
        cleaned = (
            series.fillna("")
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)
    
    def normalize_plant_series(self, series: pd.Series) -> pd.Series:
        """공장명 정규화 (비교용)"""
        normalized = series.fillna("").astype(str)
        
        # 1단계: 공백 제거
        normalized = normalized.str.replace(r"\s+", "", regex=True)
        
        # 2단계: null 토큰 처리
        lower = normalized.str.lower()
        normalized = normalized.mask(
            lower.isin(self.config.null_tokens) | (normalized == ""), 
            ""
        )
        
        # 3단계: 특수값 치환
        normalized = normalized.replace({"전사공통": "본사"})
        
        # 4단계: '공장' 접미사 제거 후 재통일
        # 예: "남양주공장" → "남양주" → "남양주공장" (통일)
        # 예: "남양주" → "남양주" → "남양주공장" (통일)
        normalized = normalized.str.replace("공장$", "", regex=True)
        
        # 5단계: 본사가 아닌 경우 '공장' 추가
        mask = (normalized != "") & (normalized != "본사")
        normalized = normalized.where(~mask, normalized + "공장")
        
        return normalized
    
    @staticmethod
    def normalize_header_cell(value: Any) -> str:
        """헤더 셀 정규화 (탐지용)"""
        if pd.isna(value):
            return ""
        
        text = str(value)
        text = text.replace("\n", " ").replace("\r", " ")
        text = re.sub(r"\s+", "", text)
        text = text.replace(":", "").replace("·", "").replace("-", "")
        
        return text.strip().lower()


# ============================================================================
# ERP 파일 처리
# ============================================================================

class ERPFileHandler:
    """ERP(종합정보시스템) 파일 처리"""
    
    REQUIRED_COLUMNS = ["공장", "자코드", "계정코드"]
    
    def __init__(self, config: Config, normalizer: DataNormalizer):
        self.config = config
        self.normalizer = normalizer
    
    def pick_latest_file(self) -> Path:
        """최신 ERP 파일 선택"""
        files = sorted(
            self.config.erp_dir.glob("*.xlsx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True
        )
        
        if not files:
            raise FileNotFoundError(
                f"ERP 파일을 찾을 수 없습니다: {self.config.erp_dir}"
            )
        
        logger.info(f"ERP 파일 선택: {files[0].name}")
        return files[0]
    
    def load_file(self, file_path: Path) -> pd.DataFrame:
        """ERP 파일 로드 및 전처리"""
        logger.info(f"ERP 파일 로딩 중: {file_path}")
        
        try:
            df = pd.read_excel(file_path, dtype=str, engine="openpyxl")
        except Exception as e:
            raise ValueError(f"ERP 파일 로딩 실패: {e}") from e
        
        # 필수 컬럼 검증
        missing_cols = set(self.REQUIRED_COLUMNS) - set(df.columns)
        if missing_cols:
            raise ValueError(
                f"ERP 파일에 필수 컬럼이 없습니다: {missing_cols}\n"
                f"현재 컬럼: {list(df.columns)}"
            )
        
        # 데이터 정규화
        df["공장"] = self.normalizer.normalize_plant_series(df["공장"])
        df["자코드"] = self.normalizer.normalize_code_series(df["자코드"])
        df["계정코드"] = self.normalizer.normalize_code_series(df["계정코드"])
        
        # 빈 자코드 제거
        df = df[df["자코드"] != ""].copy()
        
        # 0~150번 코드의 기록용 등록(본사, 금액0) 필터링
        # - ERP에는 "기록용"으로 본사에 금액 0원으로 등록된 행이 있을 수 있음
        # - 실제 투자이력 기준으로는 미등록 처리되어야 하므로, 아래 조건을 만족하는 행만 제외
        #   (공장=본사) AND (자코드 번호 0~150) AND (승인/실적/계획 금액이 모두 0)
        #
        # 주의: 본사에서 실제로 집행한 투자(금액>0)는 제외하면 안 됨.
        def is_placeholder_record(row) -> bool:
            if row.get("공장", "") != "본사":
                return False

            code = row.get("자코드", "")
            match = re.search(r'(\d{3})$', str(code))
            if not match:
                return False

            code_num = int(match.group(1))
            if not (0 <= code_num <= 150):
                return False

            # 금액 컬럼은 파일마다 다를 수 있어, 존재하는 컬럼만 사용
            amount_cols = [c for c in ["품의승인금액", "완료보고실적"] if c in row.index]
            if not amount_cols:
                # 금액 컬럼이 없으면 "0원 기록용" 여부를 확인할 수 없음 → 제거하지 않음
                return False

            # 문자열/콤마 포함 가능하므로 숫자 변환
            def to_num(v) -> float:
                try:
                    s = "" if pd.isna(v) else str(v)
                    s = s.replace(",", "").strip()
                    return float(s) if s else 0.0
                except Exception:
                    return 0.0

            return all(to_num(row[c]) == 0.0 for c in amount_cols)

        before_count = len(df)

        df["_is_placeholder"] = df.apply(is_placeholder_record, axis=1)
        df = df[~df["_is_placeholder"]].copy()
        df.drop(columns=["_is_placeholder"], inplace=True)

        placeholder_removed = before_count - len(df)
        if placeholder_removed > 0:
            logger.info(f"기록용 등록(본사, 0~150번, 금액0) 제거: {placeholder_removed}건")
        
        

        # =========================================================
        # ERP 입력오류 후보 검출(제외는 validate 단계에서 수행)
        # - "본사 단독 + 금액 입력" 형태의 코드 목록을 후보로 뽑아 attrs에 저장
        # =========================================================
        amt_cols = [c for c in ["품의승인금액", "완료보고실적"] if c in df.columns]
        if amt_cols:
            tmp = df.copy()
            for c in amt_cols:
                tmp[c] = pd.to_numeric(tmp[c].astype(str).str.replace(",", "", regex=False).str.strip(),
                                       errors="coerce").fillna(0.0)

            g = tmp.groupby("자코드", dropna=False)
            has_hq = g["공장"].apply(lambda s: (s == "본사").any())
            has_non_hq = g["공장"].apply(lambda s: (s != "본사").any())

            hq_amt_sum = tmp[tmp["공장"] == "본사"].groupby("자코드")[amt_cols].sum().sum(axis=1)
            hq_amt_sum = hq_amt_sum.reindex(has_hq.index).fillna(0.0)

            bad_mask = (has_hq) & (~has_non_hq) & (hq_amt_sum > 0)
            cand_codes = sorted(set(bad_mask[bad_mask].index.astype(str)))

            df.attrs["hq_amount_error_candidate_codes"] = cand_codes
            if cand_codes:
                df.attrs["hq_amount_error_candidate_df"] = df[df["자코드"].isin(cand_codes)].copy()
                logger.warning(f"ERP 입력오류 후보(본사 단독+금액입력) 코드 {len(cand_codes)}개 검출 (최종 제외여부는 실제 공장 기준으로 판단)")
            else:
                df.attrs["hq_amount_error_candidate_df"] = pd.DataFrame(columns=df.columns)
        else:
            df.attrs["hq_amount_error_candidate_codes"] = []
            df.attrs["hq_amount_error_candidate_df"] = pd.DataFrame(columns=df.columns)
        logger.info(f"ERP 데이터 로딩 완료: {len(df)}행")
        return df
    
    def aggregate_by_key(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        자코드 + 공장 단위로 집계
        
        0~150번 코드는 공장별로 별도 등록되므로 공장 구분 필요
        151번 이상 자코드는 전사 고유값이지만, 공장 정보도 함께 집계
        """
        # 코드 번호 추출
        def get_code_number(code):
            match = re.search(r'(\d{3})$', code)
            return int(match.group(1)) if match else -1
        
        df["코드번호"] = df["자코드"].apply(get_code_number)
        
        # 0~150번 코드: 공장별로 집계 (공장 구분 필요)
        # 151번 이상: 전사 고유값이지만 공장 정보도 저장
        aggregated = (
            df.assign(계정코드입력여부=df["계정코드"].ne(""))
            .groupby(["자코드", "공장"], dropna=False)
            .agg(
                ERP행수=("자코드", "size"),
                계정코드입력여부=("계정코드입력여부", "any"),
                코드번호=("코드번호", "first")
            )
            .reset_index()
        )
        
        logger.info(
            f"ERP 집계 완료: {len(aggregated)}개 (자코드+공장 조합), "
            f"고유 자코드: {df['자코드'].nunique()}개"
        )
        return aggregated


# ============================================================================
# 실제 투자실적 파일 처리
# ============================================================================

class ActualFileHandler:
    """실제 투자실적 파일 처리"""
    
    def __init__(self, config: Config, normalizer: DataNormalizer):
        self.config = config
        self.normalizer = normalizer
    
    def pick_file_with_dialog(self, initial_dir: Path | None = None) -> Path:
        """GUI 다이얼로그로 파일 선택"""
        if not TKINTER_AVAILABLE:
            raise RuntimeError(
                "tkinter를 사용할 수 없는 환경입니다. "
                "파일 경로를 직접 문자열로 전달하세요."
            )
        
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        
        file_path = filedialog.askopenfilename(
            title="실제 투자실적 엑셀 파일을 선택하세요",
            initialdir=str(initial_dir) if initial_dir else None,
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls"), 
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            messagebox.showinfo("취소", "파일 선택이 취소되었습니다.")
            raise SystemExit(0)
        
        logger.info(f"실제 파일 선택: {Path(file_path).name}")
        return Path(file_path)
    
    def find_header_row(
        self, 
        excel_path: Path, 
        sheet_name: str | int = 0
    ) -> int:
        """
        헤더 행 탐지
        - 2행 헤더/병합셀 고려
        - 키워드 포함 여부로 판단
        """
        preview = pd.read_excel(
            excel_path,
            sheet_name=sheet_name,
            header=None,
            nrows=self.config.max_scan_rows,
            dtype=str,
            engine="openpyxl",
        )
        
        def count_keyword_hits(cells_normalized: list[str]) -> set[str]:
            """정규화된 셀 목록에서 키워드 매칭 개수"""
            hits = set()
            for key, keywords in self.config.column_aliases.items():
                if any(
                    any(kw in cell for kw in keywords)
                    for cell in cells_normalized if cell
                ):
                    hits.add(key)
            return hits
        
        n = len(preview)
        for i in range(max(0, n - 1)):
            # i행 정규화
            row1 = [
                self.normalizer.normalize_header_cell(v) 
                for v in preview.iloc[i].tolist()
            ]
            
            # i+1행 정규화
            row2 = (
                [
                    self.normalizer.normalize_header_cell(v)
                    for v in preview.iloc[i + 1].tolist()
                ]
                if i + 1 < n else []
            )
            
            # 병합셀 대응: 두 행 결합
            combined = [a + b for a, b in zip(row1, row2)] if row2 else []
            
            # 키워드 매칭
            hits = (
                count_keyword_hits(row1) | 
                count_keyword_hits(row2) | 
                count_keyword_hits(combined)
            )
            
            if len(hits) >= self.config.min_header_hits:
                logger.info(f"헤더 행 발견: {i}행 (매칭: {hits})")
                return i
        
        raise ValueError(
            "헤더 행을 찾지 못했습니다. "
            "2행 헤더/병합셀/헤더명 상이 가능성을 확인하세요."
        )
    
    @staticmethod
    def flatten_multiindex_columns(df: pd.DataFrame) -> pd.DataFrame:
        """MultiIndex 컬럼을 단일 레벨로 평탄화"""
        if isinstance(df.columns, pd.MultiIndex):
            new_cols = []
            for tup in df.columns:
                parts = [
                    str(p).strip() 
                    for p in tup 
                    if p is not None and str(p).strip().lower() != "nan"
                ]
                new_cols.append("".join(parts))
            df = df.copy()
            df.columns = new_cols
        else:
            df.columns = [str(c).strip() for c in df.columns]
        
        return df
    
    @staticmethod
    def find_column_by_keywords(
        columns: list[str], 
        *keywords: str
    ) -> str:
        """키워드를 모두 포함하는 컬럼 찾기"""
        for col in columns:
            if all(kw in str(col) for kw in keywords):
                return col
        
        raise ValueError(
            f"필수 컬럼을 찾지 못했습니다: keywords={keywords}, "
            f"현재 컬럼={columns}"
        )
    
    def load_file(
        self, 
        file_path: Path, 
        sheet_name: str | int = 0
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        """
        실제 투자실적 파일 로드
        
        Returns:
            (등록대상 데이터, 전체 데이터)
        """
        logger.info(f"실제 파일 로딩 중: {file_path}")
        
        xls = pd.ExcelFile(file_path, engine="openpyxl")
        
        # 시트 후보 목록 생성
        candidates = self._get_sheet_candidates(xls, sheet_name)
        
        # 각 시트에서 헤더 탐지 및 로딩 시도
        df, used_sheet, header_row = self._try_load_sheets(
            file_path, 
            candidates
        )
        
        # 필수 컬럼 매핑
        mapped_df = self._map_columns(df)
        
        # 데이터 정규화
        normalized_df = self._normalize_actual_data(mapped_df)
        
        # 등록 대상 필터링
        targets = self._filter_targets(normalized_df)
        
        # 메타데이터 저장
        targets.attrs["used_sheet"] = used_sheet
        targets.attrs["header_row"] = header_row
        
        logger.info(
            f"실제 데이터 로딩 완료: 전체 {len(normalized_df)}행, "
            f"등록대상 {len(targets)}행"
        )
        
        return targets, normalized_df
    
    def _get_sheet_candidates(
        self, 
        xls: pd.ExcelFile, 
        sheet_name: str | int
    ) -> list[str]:
        """시트 후보 목록 생성"""
        candidates = []
        
        if isinstance(sheet_name, int):
            if not 0 <= sheet_name < len(xls.sheet_names):
                raise ValueError(
                    f"시트 인덱스 범위 초과: {sheet_name}, "
                    f"최대: {len(xls.sheet_names) - 1}"
                )
            candidates.append(xls.sheet_names[sheet_name])
        else:
            candidates.append(str(sheet_name))
        
        # 나머지 시트 추가 (fallback)
        for sn in xls.sheet_names:
            if sn not in candidates:
                candidates.append(sn)
        
        return candidates
    
    def _try_load_sheets(
        self, 
        file_path: Path, 
        candidates: list[str]
    ) -> tuple[pd.DataFrame, str, int]:
        """여러 시트에서 로딩 시도"""
        last_error = None
        
        for sheet_name in candidates:
            try:
                header_row = self.find_header_row(file_path, sheet_name)
                
                # 2행 헤더 우선 시도
                try:
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet_name,
                        header=[header_row, header_row + 1],
                        dtype=str,
                        engine="openpyxl"
                    )
                    df = self.flatten_multiindex_columns(df)
                except Exception:
                    # 1행 헤더 fallback
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet_name,
                        header=header_row,
                        dtype=str,
                        engine="openpyxl"
                    )
                    df = self.flatten_multiindex_columns(df)
                
                logger.info(f"시트 로딩 성공: '{sheet_name}' (헤더: {header_row}행)")
                return df, sheet_name, header_row
                
            except Exception as e:
                last_error = e
                logger.warning(f"시트 '{sheet_name}' 로딩 실패: {e}")
                continue
        
        raise ValueError(
            f"모든 시트에서 로딩 실패. 마지막 오류: {last_error}"
        )
    
    def _map_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """컬럼명 매핑"""
        cols = list(df.columns)
        
        # 기본 컬럼 매핑
        plant_col = self.find_column_by_keywords(cols, "공장")
        dept_col = self.find_column_by_keywords(cols, "부서")
        name_col = self.find_column_by_keywords(cols, "투자명")
        capex_col = self.find_column_by_keywords(cols, "자본적")
        opex_col = self.find_column_by_keywords(cols, "수익적")
        
        # 투자코드 컬럼 처리 (모코드/자코드)
        parent_col, child_col = self._identify_code_columns(cols)
        
        # 필요한 컬럼만 선택
        selected_cols = [
            plant_col, dept_col, parent_col, 
            name_col, child_col, capex_col, opex_col
        ]
        
        result = df[selected_cols].copy()
        
        # 표준 컬럼명으로 변경 (C열=모코드, E열=자코드)
        result.rename(columns={
            plant_col: "공장",
            dept_col: "부서",
            parent_col: "모코드_원본(C열_투자코드)",
            name_col: "투자명",
            child_col: "자코드(E열_투자코드)",
            capex_col: "자본적",
            opex_col: "수익적",
        }, inplace=True)
        
        return result
    
    @staticmethod
    def _identify_code_columns(columns: list[str]) -> tuple[str, str]:
        """모코드/자코드 컬럼 식별"""
        invest_code_cols = [
            c for c in columns if "투자코드" in str(c)
        ]
        
        if len(invest_code_cols) >= 2:
            # 첫 번째를 모코드, 마지막을 자코드로
            return invest_code_cols[0], invest_code_cols[-1]
        
        if len(invest_code_cols) == 1:
            # 모코드/자코드 명시적 컬럼 찾기
            parent_cols = [
                c for c in columns 
                if "모코드" in str(c) or "parent" in str(c).lower()
            ]
            child_cols = [
                c for c in columns 
                if "자코드" in str(c) or "child" in str(c).lower()
            ]
            
            if parent_cols and child_cols:
                return parent_cols[0], child_cols[0]
        
        raise ValueError(
            f"투자코드 컬럼(모코드/자코드)을 식별할 수 없습니다. "
            f"발견된 컬럼: {invest_code_cols}"
        )
    
    def _normalize_actual_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """실제 데이터 정규화"""
        result = df.copy()
        
        # 공장/코드 정규화
        result["공장"] = self.normalizer.normalize_plant_series(result["공장"])
        result["자코드(E열_투자코드)"] = self.normalizer.normalize_code_series(
            result["자코드(E열_투자코드)"]
        )
        result["모코드_원본(C열_투자코드)"] = self.normalizer.normalize_code_series(
            result["모코드_원본(C열_투자코드)"]
        )
        
        # 금액 정규화
        result["자본적"] = self.normalizer.to_number_series(result["자본적"])
        result["수익적"] = self.normalizer.to_number_series(result["수익적"])
        
        # 승인금액 합계
        result["승인금액합(C+D)"] = result["자본적"] + result["수익적"]
        
        # 공장 Forward Fill: 공장명이 병합되어 있을 수 있으므로
        # 빈 공장은 위 행의 공장 사용
        result["공장_산출"] = (
            result["공장"]
            .replace(["", "-"], pd.NA)
            .ffill()
            .fillna("")
        )
        # 공장_산출을 공장으로 덮어쓰기
        result["공장"] = result["공장_산출"]
        result.drop(columns=["공장_산출"], inplace=True)
        
        # 자코드 Forward Fill: 투자코드가 없는 행은 위 행의 코드 사용
        # "-"(투자계획)은 제외하고 forward fill
        # ---------------------------------------------------------
        # 자코드 산출 (중요 수정)
        # - 코드 상속(ffill)은 세부행(└ ...)에만 허용
        # - 헤더/구분행(예: ". xxx", 공장 "-" 등)에는 코드가 전파되지 않게 막음
        # ---------------------------------------------------------
        raw_code = result["자코드(E열_투자코드)"].replace(["", "-"], pd.NA)

        name_s = result["투자명"].fillna("").astype(str)
        is_detail = name_s.str.contains("└", regex=False)

        base_code = raw_code.where(raw_code.notna(), pd.NA)
        ffilled_code = base_code.ffill()

        result["자코드_산출"] = ""

        # 1) 원본 코드가 있는 행은 그대로 사용
        result.loc[raw_code.notna(), "자코드_산출"] = raw_code.loc[raw_code.notna()].astype(str)

        # 2) 원본 코드가 없더라도 "세부행"이면 직전 코드 상속
        mask_inherit = raw_code.isna() & is_detail
        result.loc[mask_inherit, "자코드_산출"] = ffilled_code.loc[mask_inherit].fillna("").astype(str)

        # 공장값이 '-' / 공란인 행은 세부행이 아닌 이상 코드 전파 금지(구분행 방지)
        bad_plant = result["공장"].fillna("").astype(str).str.strip().isin(["", "-"])
        result.loc[bad_plant & (~is_detail) & raw_code.isna(), "자코드_산출"] = ""

        
        # 모코드 산출: 자코드 번호가 151 이상이면 C열 모코드 사용, 아니면 공란
        def extract_parent_code(row):
            child_code = row["자코드_산출"]
            if not child_code or child_code == "-":
                return ""
            
            # 코드에서 숫자 부분 추출
            match = re.search(r'(\d{3})$', child_code)
            if match:
                code_num = int(match.group(1))
                # 151 이상이면 모코드 컬럼에서 가져오기
                if code_num >= 151:
                    # 모코드도 forward fill 적용
                    parent = row["모코드_원본(C열_투자코드)"]
                    return parent if parent and parent != "-" else ""
            
            return ""
        
        # 임시로 모코드도 forward fill
        result["모코드_ffill_temp"] = (
            result["모코드_원본(C열_투자코드)"]
            .replace(["", "-"], pd.NA)
            .ffill()
            .fillna("")
        )
        
        # 모코드 산출
        result["모코드_산출"] = result.apply(
            lambda row: extract_parent_code({
                "자코드_산출": row["자코드_산출"],
                "모코드_원본(C열_투자코드)": row["모코드_ffill_temp"]
            }),
            axis=1
        )
        
        # 임시 컬럼 제거
        result.drop(columns=["모코드_ffill_temp"], inplace=True)

        # =========================================================
        # (추가) 자코드 중복 방지: 자코드 기준 상세행 합산 + 투자명 forward fill
        # - 실제 파일에서 한 투자코드 아래에 세부 공사항목이 여러 줄로 존재할 수 있음
        # - 투자코드는 모코드가 같더라도 자코드는 중복될 수 없으므로,
        #   동일 자코드로 내려오는 세부 행은 금액을 합산하여 1행으로 통합한다.
        # - 투자명은 "투자코드가 있는 행"의 투자명을 기준으로 forward fill한다.
        # =========================================================

        # 부서 Forward Fill (병합셀/빈칸 대응)
        result["부서_산출"] = (
            result["부서"]
            .replace(["", "-"], pd.NA)
            .ffill()
            .fillna("")
        )
        result["부서"] = result["부서_산출"]

        # 투자명 Forward Fill: 원본 자코드 셀(E열)이 존재하는 행의 투자명을 기준으로 ffill
        code_cell = result["자코드(E열_투자코드)"].replace(["", "-"], pd.NA)
        name_base = result["투자명"].where(code_cell.notna(), pd.NA)
        result["투자명"] = name_base.ffill().fillna("")

        # 동일 자코드(+공장) 상세행 합산
        group_cols = ["공장", "부서", "자코드_산출", "모코드_산출"]
        aggregated = (
            result
            .groupby(group_cols, dropna=False, as_index=False)
            .agg(
                투자명=("투자명", "first"),
                **{
                    "모코드_원본(C열_투자코드)": ("모코드_원본(C열_투자코드)", "first"),
                    "자코드(E열_투자코드)": ("자코드(E열_투자코드)", "first"),
                    "자본적": ("자본적", "sum"),
                    "수익적": ("수익적", "sum"),
                    "승인금액합(C+D)": ("승인금액합(C+D)", "sum"),
                }
            )
        )

        # 자코드(E열)는 실제로는 그룹 키(자코드_산출)가 맞으므로 통일
        aggregated["자코드(E열_투자코드)"] = aggregated["자코드_산출"]

        # 반환 컬럼 정리(이후 단계에서 필요한 컬럼 위주)
        keep_cols = [
            "공장", "부서", "투자명",
            "모코드_원본(C열_투자코드)", "자코드(E열_투자코드)",
            "자본적", "수익적", "승인금액합(C+D)",
            "자코드_산출", "모코드_산출",
        ]
        aggregated = aggregated[keep_cols].copy()

        logger.info(
            f"실제 데이터 상세행 합산 완료: {len(result)}행 → {len(aggregated)}행 "
            f"(고유 자코드: {aggregated['자코드_산출'].nunique()}개)"
        )

        return aggregated
    
    @staticmethod
    def _is_valid_investment_code(code: str) -> bool:
        """
        유효한 투자코드 형식 확인 (알파벳2자 + 숫자3자)
        
        Args:
            code: 투자코드 문자열
        
        Returns:
            유효한 형식이면 True
        """
        if not code or code == "-" or code == "":
            return False
        
        # 알파벳 2자 + 숫자 3자 패턴 (예: AB123, CD050)
        pattern = r'^[A-Za-z]{2}\d{3}$'
        return bool(re.match(pattern, code))
    
    @staticmethod
    def _get_code_number(code: str) -> int:
        """
        투자코드에서 숫자 부분 추출 + 0~150 모코드(XX000~XX150) 구분용 보정

        - 0~150 규칙은 사용자가 정의한 '모코드(XX000~XX150)' 체계에만 적용.
        - 그 외 접두어(예: ZS028)는 숫자만 0~150이라도 모코드 체계로 보지 않고,
          공장 불일치 시 '자코드만으로 재매칭' 대상(151+)으로 처리한다.

        Args:
            code: 투자코드 (예: XX010, ZS010, ZS163)

        Returns:
            - 접두어가 XX인 경우: 끝 3자리 숫자(int)
            - 그 외 접두어: 9999 (151+로 취급)
            - 실패시: -1
        """
        if code is None:
            return -1
        s = str(code).strip()
        if s == "":
            return -1

        m = re.match(r'^([A-Za-z]+)(\d{3})$', s)
        if not m:
            m2 = re.search(r'(\d{3})$', s)
            return int(m2.group(1)) if m2 else -1

        prefix = m.group(1).upper()
        num = int(m.group(2))

        if prefix == "XX":
            return num
        return 9999
    
    @staticmethod
    def _filter_targets(df: pd.DataFrame) -> pd.DataFrame:
        """
        등록 대상 필터링
        
        필터링 기준:
        1. 투자계획("-")이 아닌 실제 투자코드(XX000~XX999 형식)만
        2. 해당 행에 금액이 있는 항목만
        3. 자코드_산출 컬럼 사용 (forward fill된 코드)
        4. 합산 행(소계, 공장, 전사 등) 제외
        
        투자 구조:
        - 투자계획 (코드: "-") → ERP 미등록 대상 (제외)
        - 0~150번 코드 (E열에 직접 입력, C열 공란) → ERP 등록 대상 ✓
        - 151번 이상 코드 (E열에 입력, C열에 모코드) → ERP 등록 대상 ✓
        - 코드 없고 금액만 있는 행 → 위 행의 코드 사용 (forward fill) ✓
        """
        result = df.copy()
        
        # 유효한 투자코드 판별 함수
        def is_valid_code(code: str) -> bool:
            return ActualFileHandler._is_valid_investment_code(code)
        
        # 합산 행 판별 함수
        def is_summary_row(row) -> bool:
            """
            소계/합계/구분 행 판별 (실제파일의 요약행 제거)
            - 투자명에 소계/합계/총계/전사 등 키워드가 있으면 요약행으로 판단
            - 공장/부서가 '-' 로 채워진 행은 강하게 요약행 후보
            - 투자코드가 비어있고(또는 '-') 금액만 존재하는 합계행 패턴도 제거
            """
            투자명 = str(row.get("투자명", "")).strip().lower()
            공장 = str(row.get("공장", "")).strip()
            부서 = str(row.get("부서", "")).strip()

            summary_keywords = ["소계", "합계", "총계", "전사", "계"]

            if any(kw in 투자명 for kw in summary_keywords):
                return True

            # 공장/부서가 '-' 또는 공란이고, 투자코드가 비어있으면(또는 '-') 요약행으로 판단
            code_raw = str(row.get("자코드(E열_투자코드)", "")).strip()
            if (공장 in {"-", ""} and 부서 in {"-", ""}) and (code_raw in {"-", "", "nan"}):
                return True

            # 부서가 '-' 또는 공란이고 투자명 길이가 짧으면 요약행으로 판단 (기존 로직 보강)
            if (부서 in {"-", ""}) and (len(투자명) <= 4):
                return True

            return False
        
        # 자코드는 이미 forward fill된 "자코드_산출" 사용
        result["자코드"] = result["자코드_산출"]
        
        # 합산 행 제거
        result["_is_summary"] = result.apply(is_summary_row, axis=1)
        summary_count = result["_is_summary"].sum()
        
        if summary_count > 0:
            logger.info(f"합산 행(소계/합계) 제거: {summary_count}건")
        
        # 필터링: 유효한 자코드가 있고 금액이 있고 합산 행이 아닌 것만
        targets = result[
            (result["자코드"] != "") &
            (result["자코드"] != "-") &
            (result["승인금액합(C+D)"] > 0) &
            (~result["_is_summary"])  # 합산 행 제외
        ].copy()
        
        # 임시 컬럼 제거
        targets.drop(columns=["_is_summary"], inplace=True, errors='ignore')
        
        # 코드 분류 플래그 추가
        targets["코드번호"] = targets["자코드"].apply(
            ActualFileHandler._get_code_number
        )
        
        targets["모코드통째사용"] = (
            (targets["코드번호"] >= 0) & 
            (targets["코드번호"] <= 150) &
            (targets["모코드_산출"] == "")
        )
        
        targets["자코드사용"] = targets["코드번호"] >= 151
        
        targets["코드없이금액만"] = (
            (targets["자코드(E열_투자코드)"] == "") | 
            (targets["자코드(E열_투자코드)"] == "-")
        ) & (targets["자코드"] != "")
        
        # 통계 정보
        total_targets = len(targets)
        parent_only = targets["모코드통째사용"].sum()
        child_code = targets["자코드사용"].sum()
        no_code_but_amount = targets["코드없이금액만"].sum()
        
        logger.info(
            f"등록 대상 필터링 완료: "
            f"총 {len(df)}행 중 {total_targets}행 추출 "
            f"(0~150번: {parent_only}건, 151번↑: {child_code}건, "
            f"코드없이금액: {no_code_but_amount}건)"
        )
        
        return targets


# ============================================================================
# 검증 로직
# ============================================================================

class InvestmentValidator:
    """투자실적 검증 담당 클래스"""
    
    def __init__(self, config: Config):
        self.config = config
        self.normalizer = DataNormalizer(config)
        self.erp_handler = ERPFileHandler(config, self.normalizer)
        self.actual_handler = ActualFileHandler(config, self.normalizer)
    
    def validate(
        self, 
        actual_path: str | Path, 
        actual_sheet_name: str | int = 0
    ) -> dict[str, Any]:
        """
        투자실적 검증 수행
        
        Returns:
            검증 결과 딕셔너리
        """
        # ERP 데이터 로드
        erp_path = self.erp_handler.pick_latest_file()
        erp_df = self.erp_handler.load_file(erp_path)

        # 실제 데이터 로드
        targets, actual_all = self.actual_handler.load_file(
            Path(actual_path), 
            actual_sheet_name
        )

        # =========================================================
        # ERP 입력오류 최종 판정: "본사 단독 + 금액 입력" 후보 중
        # 실제 투자이력(등록대상)에서 '생산 공장'으로 등장하는 코드만 오류로 확정
        #
        # - 실제 파일이 본사/전사공통으로 들어오는 케이스가 있어,
        #   실제 공장 기준 판단은 targets의 공장 값을 활용.
        # - 단, 연구소공장 등 예외를 피하기 위해 '연구소' 포함 공장은 제외.
        #   (현 시스템에서 연구소 투자 등록 규칙이 다를 수 있어 안전측)
        # =========================================================
        cand_codes = set(erp_df.attrs.get("hq_amount_error_candidate_codes", []))
        final_bad_codes = set()
        if cand_codes:
            # targets는 실제 등록대상(요약행 제거/필터 적용 후)
            prod_mask = (
                (targets["공장"] != "본사") &
                (targets["공장"] != "-") &
                (targets["공장"].astype(str).str.strip() != "") &
                (~targets["공장"].astype(str).str.contains("연구소", na=False))
            )
            prod_codes = set(targets.loc[prod_mask, "자코드"].astype(str))
            final_bad_codes = cand_codes & prod_codes

        if final_bad_codes:
            before = len(erp_df)
            erp_df = erp_df[~erp_df["자코드"].isin(final_bad_codes)].copy()
            removed = before - len(erp_df)
            logger.warning(f"ERP 입력오류 확정(본사 단독+금액입력): 코드 {len(final_bad_codes)}개, 행 {removed}건 제외 → 전산미등록으로 처리")
            # 리포트용 DF 보관
            erp_df.attrs["hq_amount_error_df"] = erp_df.attrs.get("hq_amount_error_candidate_df", pd.DataFrame()).copy()
            erp_df.attrs["hq_amount_error_codes"] = sorted(final_bad_codes)
        else:
            erp_df.attrs["hq_amount_error_df"] = pd.DataFrame(columns=erp_df.columns)
            erp_df.attrs["hq_amount_error_codes"] = []

        # 집계는 최종 ERP(필터 후) 기준으로 수행
        erp_aggregated = self.erp_handler.aggregate_by_key(erp_df)

        # 병합 로직 개선:
        # 1) 먼저 자코드 + 공장으로 정확히 매칭 시도
        # 2) 매칭 실패 시, 151번 이상 코드는 자코드만으로 재매칭 (전사 고유값)
        
        # Step 1: 자코드 + 공장으로 병합
        merged = targets.merge(
            erp_aggregated,
            on=["자코드", "공장"],
            how="left",
            suffixes=('', '_ERP')
        )
        
        # Step 2: 151번 이상 코드 중 매칭 실패한 경우, 자코드만으로 재시도
        unmatched_151plus = merged[
            (merged["ERP행수"].isna()) & 
            (merged["코드번호"] >= 151)
        ].copy()
        
        if len(unmatched_151plus) > 0:
            logger.info(
                f"151번 이상 코드 중 공장 불일치 {len(unmatched_151plus)}건 → "
                f"자코드만으로 재매칭 시도"
            )
            
            # 자코드만으로 매칭 시도
            remerged = unmatched_151plus[["자코드"]].merge(
                erp_aggregated[["자코드", "ERP행수", "계정코드입력여부"]].drop_duplicates(subset=["자코드"]),
                on=["자코드"],
                how="left",
                suffixes=('', '_retry')
            )
            
            # 재매칭 성공한 행 업데이트
            for idx in unmatched_151plus.index:
                자코드 = unmatched_151plus.loc[idx, "자코드"]
                match_result = remerged[remerged["자코드"] == 자코드]
                
                if len(match_result) > 0 and not pd.isna(match_result.iloc[0]["ERP행수"]):
                    merged.loc[idx, "ERP행수"] = match_result.iloc[0]["ERP행수"]
                    merged.loc[idx, "계정코드입력여부"] = match_result.iloc[0]["계정코드입력여부"]
            
            rematched = merged.loc[unmatched_151plus.index, "ERP행수"].notna().sum()
            if rematched > 0:
                logger.info(f"재매칭 성공: {rematched}건")
        
        # 검증 1: 전산 미등록
        missing_in_erp = merged[merged["ERP행수"].isna()].copy()
        
        # 검증 2: 계정코드 미입력
        account_missing = merged[
            (merged["ERP행수"].notna()) & 
            (merged["계정코드입력여부"] == False)
        ].copy()
        
        logger.info(
            f"검증 완료 - 전산미등록: {len(missing_in_erp)}건, "
            f"계정코드미입력: {len(account_missing)}건"
        )
        
        return {
            "erp_path": erp_path,
            "actual_path": Path(actual_path),
            "erp_data": erp_df,
            "erp_aggregated": erp_aggregated,
            "targets": targets,
            "actual_all": actual_all,
            "merged": merged,
            "missing_in_erp": missing_in_erp,
            "account_missing": account_missing,
        }
    
    def generate_report(
        self, 
        validation_result: dict[str, Any]
    ) -> Path:
        """검증 결과 리포트 생성 (Excel)"""
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = self.config.output_dir / f"투자실적_점검결과_{timestamp}.xlsx"
        
        # 요약 정보
        targets = validation_result["targets"]
        summary = pd.DataFrame([{
            "종합정보시스템 파일": validation_result["erp_path"].name,
            "실제 투자실적 파일": validation_result["actual_path"].name,
            "실제 등록대상 행수": int(len(targets)),
            "실제 등록대상 자코드수": int(targets["자코드"].nunique()),
            "종합정보 자코드수": int(
                validation_result["erp_data"]["자코드"].nunique()
            ),
            "전산미등록": int(len(validation_result["missing_in_erp"])),
            "계정코드미입력": int(len(validation_result["account_missing"])),
            "실제파일_사용시트": targets.attrs.get("used_sheet", ""),
            "실제파일_헤더행(0기준)": targets.attrs.get("header_row", ""),
        }])
        
        # Excel 파일 작성
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary.to_excel(writer, index=False, sheet_name="요약")
            
            validation_result["missing_in_erp"].to_excel(
                writer, index=False, sheet_name="전산미등록"
            )
            
            validation_result["account_missing"].to_excel(
                writer, index=False, sheet_name="계정코드미입력"
            )
            
            validation_result["erp_aggregated"].to_excel(
                writer, index=False, sheet_name="종합정보 집계(자코드+공장)"
            )

            # ERP 입력오류(본사 단독 + 금액 입력) 목록
            erp_data = validation_result.get("erp_data")
            hq_err_df = None
            if erp_data is not None:
                hq_err_df = erp_data.attrs.get("hq_amount_error_df")
            if isinstance(hq_err_df, pd.DataFrame) and not hq_err_df.empty:
                hq_err_df.to_excel(writer, index=False, sheet_name="ERP입력오류(본사금액)")

            validation_result["targets"].to_excel(
                writer, index=False, sheet_name="실제대상목록(필터후)"
            )
            
            validation_result["merged"].to_excel(
                writer, index=False, sheet_name="병합결과(검토용)"
            )
        
        logger.info(f"리포트 생성 완료: {output_path}")
        return output_path


# ============================================================================
# 메인 실행
# ============================================================================

def main(
    actual_path: str | Path | None = None,
    actual_sheet_name: str | int = 0,
    initial_dir: Path | None = None,
):
    """
    메인 실행 함수
    
    Args:
        actual_path: 실제 투자실적 파일 경로 (None이면 GUI 다이얼로그)
        actual_sheet_name: 시트 이름 또는 인덱스
        initial_dir: 파일 선택 다이얼로그 초기 디렉토리
    """
    config = Config()
    validator = InvestmentValidator(config)
    
    # 실제 파일 선택
    if actual_path is None:
        actual_path = validator.actual_handler.pick_file_with_dialog(
            initial_dir=initial_dir
        )
    
    # 검증 수행
    result = validator.validate(actual_path, actual_sheet_name)
    
    # 리포트 생성
    output_path = validator.generate_report(result)
    
    print(f"\n{'='*60}")
    print(f"투자실적 검증 완료")
    print(f"{'='*60}")
    print(f"결과 파일: {output_path}")
    print(f"전산미등록: {len(result['missing_in_erp'])}건")
    print(f"계정코드미입력: {len(result['account_missing'])}건")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    # 기본 동작:
    # - GUI 환경이면 파일 선택창 사용 가능
    # - 헤드리스/서버 환경에서는 actual_path를 직접 지정해서 실행하세요.
    #
    # 예)
    #   python investment_validator.py --actual "/path/to/actual.xlsx" --erp "/path/to/erp.xlsx"
    import argparse

    parser = argparse.ArgumentParser(description="투자실적 검증 자동화")
    parser.add_argument("--actual", dest="actual_path", required=False, help="실제 투자실적 파일 경로")
    parser.add_argument("--sheet", dest="actual_sheet_name", default=0, help="시트 이름 또는 인덱스 (기본 0)")
    parser.add_argument("--erp", dest="erp_path", required=False, help="ERP(종합정보) 파일 경로 (미지정 시 최신 파일 자동 선택)")
    args = parser.parse_args()

    # Config 경로는 기본값 유지. 단, --erp를 주면 해당 파일을 우선 사용
    cfg = Config()
    validator = InvestmentValidator(cfg)

    if args.actual_path is None:
        # GUI 환경이면 파일 선택 다이얼로그
        main(actual_path=None, actual_sheet_name=args.actual_sheet_name)
    else:
        # ERP 파일 강제 지정 옵션 지원
        if args.erp_path:
            erp_p = Path(args.erp_path)
            erp_df = validator.erp_handler.load_file(erp_p)
            erp_agg = validator.erp_handler.aggregate_by_key(erp_df)
            targets, actual_all = validator.actual_handler.load_file(Path(args.actual_path), args.actual_sheet_name)

            merged = targets.merge(erp_agg, on=["자코드", "공장"], how="left", suffixes=('', '_ERP'))
            missing_in_erp = merged[merged["ERP행수"].isna()].copy()
            account_missing = merged[(merged["ERP행수"].notna()) & (merged["계정코드입력여부"] == False)].copy()

            result = {
                "erp_path": erp_p,
                "actual_path": Path(args.actual_path),
                "erp_data": erp_df,
                "erp_aggregated": erp_agg,
                "targets": targets,
                "actual_all": actual_all,
                "merged": merged,
                "missing_in_erp": missing_in_erp,
                "account_missing": account_missing,
            }
            out = validator.generate_report(result)
            print(out)
        else:
            main(actual_path=Path(args.actual_path), actual_sheet_name=args.actual_sheet_name)
