"""
투자실적 검증 자동화 프로그램 (신규 양식 대응)

실제 투자이력과 전산(MIS) 상 등록된 투자이력을 비교하여
미등록 및 부서코드/계정코드 누락 사항을 검출합니다.

신규 양식 구조:
  - 투자코드: 부서(부서코드) + 유형(계정코드) + 코드(자코드) 3분할
  - 금액: '26년 계획[최종案]' + '품의금액' 2세트 (품의금액 기준 사용)
  - 헤더: 18~19행 (상단 요약 테이블 포함)
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
    MIS_dir: Path | None = None
    output_dir: Path | None = None
    
    # 헤더 탐지 설정
    max_scan_rows: int = 200
    min_header_hits: int = 3
    
    # 정규화 설정
    null_tokens: frozenset[str] = frozenset({"nan", "none", "null"})
    
    # 컬럼 별칭 (헤더 탐지용)
    column_aliases: dict[str, list[str]] = None
    
    def __post_init__(self):
        # 기본 MIS 디렉터리 자동 감지: 여러 후보 폴더명을 확인합니다.
        if self.MIS_dir is None:
            candidates = [
                self.base_dir / "종합정보 투자실적",
                self.base_dir / "투자리스트_종합정보시스템",
                self.base_dir / "투자리스트_실제",
            ]
            for c in candidates:
                if c.exists() and c.is_dir():
                    self.MIS_dir = c
                    break
            else:
                # 후보 미발견 시 기존 경로를 기본값으로 설정
                self.MIS_dir = self.base_dir / "종합정보 투자실적"

        if self.output_dir is None:
            self.output_dir = self.base_dir / "result"

        if self.column_aliases is None:
            self.column_aliases = {
                "공장": ["공장", "공장명", "사업장", "사업장명"],
                "부서": ["부서", "부서명", "사용부서", "요청부서", "부서(팀)"],
                "투자명": ["투자명", "건명", "품의명", "과제명", "내역"],
                "투자코드": ["투자코드", "코드", "MIS코드", "모코드", "자코드", "부서코드", "유형"],
                "자본적": ["자본적", "자본", "capex", "품의금액자본적"],
                "수익적": ["수익적", "수익", "opex", "품의금액수익적"],
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
        
        # 3단계: 특수값 치환 (본부공장은 본사와 동일 조직으로 취급)
        normalized = normalized.replace({"전사공통": "본사", "본부공장": "본사", "본부": "본사"})
        
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
# MIS 파일 처리
# ============================================================================

class MISFileHandler:
    """MIS(종합정보시스템) 파일 처리"""
    
    REQUIRED_COLUMNS = ["공장", "자코드", "계정코드"]
    
    def __init__(self, config: Config, normalizer: DataNormalizer):
        self.config = config
        self.normalizer = normalizer
    
    def pick_latest_file(self) -> Path:
        """최신 MIS 파일 선택 (~$ 임시파일 제외)"""
        files = sorted(
            [p for p in self.config.MIS_dir.glob("*.xlsx") if not p.name.startswith("~$")],
            key=lambda p: p.stat().st_mtime,
            reverse=True
        )
        
        if not files:
            raise FileNotFoundError(
                f"MIS 파일을 찾을 수 없습니다: {self.config.MIS_dir}"
            )
        
        logger.info(f"MIS 파일 선택: {files[0].name}")
        return files[0]
    
    def load_file(self, file_path: Path) -> pd.DataFrame:
        """MIS 파일 로드 및 전처리"""
        logger.info(f"MIS 파일 로딩 중: {file_path}")
        
        try:
            df = pd.read_excel(file_path, dtype=str, engine="openpyxl")
        except Exception as e:
            raise ValueError(f"MIS 파일 로딩 실패: {e}") from e
        
        # 필수 컬럼 검증
        missing_cols = set(self.REQUIRED_COLUMNS) - set(df.columns)
        if missing_cols:
            raise ValueError(
                f"MIS 파일에 필수 컬럼이 없습니다: {missing_cols}\n"
                f"현재 컬럼: {list(df.columns)}"
            )
        
        # 데이터 정규화
        df["공장"] = self.normalizer.normalize_plant_series(df["공장"])
        df["자코드"] = self.normalizer.normalize_code_series(df["자코드"])
        df["계정코드"] = self.normalizer.normalize_code_series(df["계정코드"])
        
        # 부서코드 정규화 (신규 양식의 투자코드.부서와 매핑용)
        if "부서코드" in df.columns:
            df["부서코드"] = self.normalizer.normalize_code_series(df["부서코드"])
        
        # 빈 자코드 제거
        df = df[df["자코드"] != ""].copy()
        
        # 0~150번 코드의 기록용 등록(본사, 금액0) 필터링
        # - MIS에는 "기록용"으로 본사에 금액 0원으로 등록된 행이 있을 수 있음
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
        # MIS 입력오류 후보 검출(제외는 validate 단계에서 수행)
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
                logger.warning(f"MIS 입력오류 후보(본사 단독+금액입력) 코드 {len(cand_codes)}개 검출 (최종 제외여부는 실제 공장 기준으로 판단)")
            else:
                df.attrs["hq_amount_error_candidate_df"] = pd.DataFrame(columns=df.columns)
        else:
            df.attrs["hq_amount_error_candidate_codes"] = []
            df.attrs["hq_amount_error_candidate_df"] = pd.DataFrame(columns=df.columns)
        logger.info(f"MIS 데이터 로딩 완료: {len(df)}행")
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

        # 금액 컬럼 숫자화 (존재 시)
        df_work = df.copy()
        for src, dst in [("투자계획금액", "_투자계획금액"), ("품의승인금액", "_품의승인금액")]:
            if src in df_work.columns:
                df_work[dst] = self.normalizer.to_number_series(df_work[src])
            else:
                df_work[dst] = 0.0
        df_work["_계정코드입력여부"] = df_work["계정코드"].ne("")

        # 0~150번 코드: 공장별로 집계 (공장 구분 필요)
        # 151번 이상: 전사 고유값이지만 공장 정보도 저장
        aggregated = (
            df_work
            .groupby(["자코드", "공장"], dropna=False)
            .agg(
                MIS행수=("자코드", "size"),
                계정코드입력여부=("_계정코드입력여부", "any"),
                계정코드값=("계정코드", lambda s: ",".join(sorted({str(x) for x in s if str(x).strip()}))),
                투자계획금액=("_투자계획금액", "sum"),
                품의승인금액=("_품의승인금액", "sum"),
                코드번호=("코드번호", "first")
            )
            .reset_index()
        )

        logger.info(
            f"MIS 집계 완료: {len(aggregated)}개 (자코드+공장 조합), "
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
            
            # i행 자체에 최소 1개 키워드가 있어야 헤더로 인정
            # (빈 행이 i+1행 결합으로 헤더 오감지되는 것 방지)
            row1_hits = count_keyword_hits(row1)
            if len(hits) >= self.config.min_header_hits and len(row1_hits) >= 1:
                logger.info(f"헤더 행 발견: {i}행 (매칭: {hits})")
                return i
        
        raise ValueError(
            "헤더 행을 찾지 못했습니다. "
            "2행 헤더/병합셀/헤더명 상이 가능성을 확인하세요."
        )
    
    @staticmethod
    def flatten_multiindex_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        MultiIndex 컬럼을 단일 레벨로 평탄화
        
        병합셀 대응: level-0 값이 NaN인 컬럼은 직전 non-NaN level-0 값을
        전파하여 고유한 이름을 생성한다.
        예: ('품의금액', '자본적') → '품의금액자본적'
            (NaN, '수익적')      → '품의금액수익적'  (직전 '품의금액' 전파)
        """
        if isinstance(df.columns, pd.MultiIndex):
            new_cols = []
            last_top = ""
            for tup in df.columns:
                top_raw = tup[0] if len(tup) > 0 else None
                bot_raw = tup[1] if len(tup) > 1 else None
                
                top_val = (
                    str(top_raw).strip()
                    if top_raw is not None and str(top_raw).strip().lower() != "nan"
                    else ""
                )
                bot_val = (
                    str(bot_raw).strip()
                    if bot_raw is not None and str(bot_raw).strip().lower() != "nan"
                    else ""
                )
                
                if top_val:
                    last_top = top_val
                    new_cols.append(top_val + bot_val)
                elif bot_val:
                    # 부모 그룹명 전파하여 중복 방지
                    new_cols.append(last_top + bot_val)
                else:
                    new_cols.append("")
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
        """
        신규 양식 컬럼명 매핑
        
        신규 양식 구조 (flatten 후):
        - 투자명, 공장, 부서(팀)
        - 투자코드부서(=부서코드), 투자코드유형(=계정코드), 투자코드코드(=자코드) 또는 유형, 코드
        - 품의금액자본적, 품의금액수익적, 품의금액計 (금액 기준)
        """
        cols = list(df.columns)
        logger.info(f"Flatten 후 컬럼 목록: {cols}")
        
        # 기본 컬럼 매핑
        plant_col = self.find_column_by_keywords(cols, "공장")
        name_col = self.find_column_by_keywords(cols, "투자명")
        
        # 부서(팀) 컬럼 - "부서(팀)" 또는 "부서" (투자코드부서와 구분 필요)
        dept_col = self._find_team_dept_column(cols)
        
        # 투자코드 3분할: 부서코드, 유형(=계정코드), 코드(=자코드)
        dept_code_col, acct_code_col, child_code_col = self._identify_code_columns(cols)
        
        # 품의금액 자본적/수익적 (계획금액이 아닌 품의금액 기준)
        capex_col, opex_col = self._find_amount_columns(cols)
        
        # 필요한 컬럼만 선택
        selected_cols = [
            plant_col, dept_col, name_col,
            dept_code_col, acct_code_col, child_code_col,
            capex_col, opex_col
        ]
        
        result = df[selected_cols].copy()
        
        # 표준 컬럼명으로 변경
        result.rename(columns={
            plant_col: "공장",
            dept_col: "부서",
            name_col: "투자명",
            dept_code_col: "부서코드_원본",
            acct_code_col: "유형_원본(계정코드)",
            child_code_col: "코드_원본(자코드)",
            capex_col: "자본적",
            opex_col: "수익적",
        }, inplace=True)
        
        return result
    
    @staticmethod
    def _find_team_dept_column(columns: list[str]) -> str:
        """
        부서(팀) 컬럼 찾기 - '투자코드부서'와 구분
        
        신규 양식에서 '부서(팀)'(조직부서)과 '투자코드부서'(부서코드)가 공존하므로
        '(팀)' 포함 또는 '투자코드' 미포함인 부서 컬럼을 우선 선택
        """
        # 1순위: "부서(팀)" 정확 매칭
        for col in columns:
            if "부서(팀)" in str(col) or "부서 (팀)" in str(col):
                return col
        
        # 2순위: "부서"를 포함하되 "투자코드"를 포함하지 않는 컬럼
        for col in columns:
            col_str = str(col)
            if "부서" in col_str and "투자코드" not in col_str and "부서코드" not in col_str:
                return col
        
        raise ValueError(
            f"부서(팀) 컬럼을 찾지 못했습니다. 현재 컬럼: {columns}"
        )
    
    @staticmethod
    def _identify_code_columns(columns: list[str]) -> tuple[str, str, str]:
        """
        투자코드 3분할 컬럼 식별: 부서코드, 유형(계정코드), 코드(자코드)
        
        신규 양식에서 투자코드 하위에 부서/유형/코드 3개 서브컬럼이 존재.
        Flatten 후 '투자코드부서', '투자코드유형' 또는 '유형', '코드' 등으로 나타남.
        
        Returns:
            (부서코드 컬럼, 유형(계정코드) 컬럼, 코드(자코드) 컬럼)
        """
        # 투자코드 관련 컬럼 찾기
        invest_cols = [c for c in columns if "투자코드" in str(c)]
        
        dept_code_col = None
        acct_code_col = None
        child_code_col = None
        
        # '투자코드부서' 찾기
        for c in columns:
            if "투자코드" in str(c) and "부서" in str(c):
                dept_code_col = c
                break
        
        # '유형' 찾기 ('투자코드유형' 또는 단독 '유형')
        for c in columns:
            cs = str(c)
            if ("유형" in cs) and ("투자" not in cs or "투자코드" in cs):
                if "투자유형" not in cs and "투자구분" not in cs:
                    acct_code_col = c
                    break
        
        # '코드' 찾기 ('투자코드코드' 또는 단독 '코드')
        for c in columns:
            cs = str(c)
            if cs.endswith("코드") and "부서코드" not in cs and "계정코드" not in cs:
                if "투자코드코드" in cs or cs == "코드":
                    child_code_col = c
                    break
        
        # fallback: invest_cols에서 위치 기반 추정
        if not all([dept_code_col, acct_code_col, child_code_col]):
            # 위치 기반: 투자코드 그룹 직후의 3개 컬럼
            if invest_cols:
                anchor_idx = columns.index(invest_cols[0])
                if dept_code_col is None and anchor_idx < len(columns):
                    dept_code_col = columns[anchor_idx]  # 투자코드부서
                if acct_code_col is None and anchor_idx + 1 < len(columns):
                    acct_code_col = columns[anchor_idx + 1]  # 유형
                if child_code_col is None and anchor_idx + 2 < len(columns):
                    child_code_col = columns[anchor_idx + 2]  # 코드
        
        if not all([dept_code_col, acct_code_col, child_code_col]):
            raise ValueError(
                f"투자코드 3분할 컬럼(부서코드/유형/코드)을 식별할 수 없습니다. "
                f"투자코드 관련 컬럼: {invest_cols}, 전체 컬럼: {columns}"
            )
        
        logger.info(
            f"투자코드 컬럼 매핑: 부서코드={dept_code_col}, "
            f"유형(계정코드)={acct_code_col}, 코드(자코드)={child_code_col}"
        )
        return dept_code_col, acct_code_col, child_code_col
    
    @staticmethod
    def _find_amount_columns(columns: list[str]) -> tuple[str, str]:
        """
        품의금액 자본적/수익적 컬럼 찾기
        
        신규 양식에서 '26년 계획'과 '품의금액' 2세트가 있으므로
        '품의금액' 하위의 자본적/수익적을 우선 선택
        """
        capex_col = None
        opex_col = None
        
        # 1순위: "품의금액자본적", "품의금액수익적" (정확 매칭)
        for col in columns:
            if "품의금액" in str(col) and "자본적" in str(col):
                capex_col = col
            if "품의금액" in str(col) and "수익적" in str(col):
                opex_col = col
        
        if capex_col and opex_col:
            return capex_col, opex_col
        
        # 2순위: "품의" + "자본적"/"수익적"
        for col in columns:
            if "품의" in str(col) and "자본" in str(col) and capex_col is None:
                capex_col = col
            if "품의" in str(col) and "수익" in str(col) and opex_col is None:
                opex_col = col
        
        if capex_col and opex_col:
            return capex_col, opex_col
        
        # 3순위: 일반 "자본적"/"수익적" (첫 번째 매칭 - 호환용)
        for col in columns:
            if "자본적" in str(col) and capex_col is None:
                capex_col = col
            if "수익적" in str(col) and opex_col is None:
                opex_col = col
        
        if capex_col and opex_col:
            return capex_col, opex_col
        
        raise ValueError(
            f"품의금액 자본적/수익적 컬럼을 찾지 못했습니다. "
            f"현재 컬럼: {columns}"
        )
    
    def _normalize_actual_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        실제 데이터 정규화 (신규 양식 대응)
        
        신규 양식 컬럼: 공장, 부서, 투자명, 부서코드_원본, 유형_원본(계정코드), 코드_원본(자코드), 자본적, 수익적
        """
        result = df.copy()
        
        # 공장/코드 정규화
        result["공장"] = self.normalizer.normalize_plant_series(result["공장"])
        result["코드_원본(자코드)"] = self.normalizer.normalize_code_series(
            result["코드_원본(자코드)"]
        )
        result["부서코드_원본"] = self.normalizer.normalize_code_series(
            result["부서코드_원본"]
        )
        result["유형_원본(계정코드)"] = self.normalizer.normalize_code_series(
            result["유형_원본(계정코드)"]
        )
        
        # 금액 정규화
        result["자본적"] = self.normalizer.to_number_series(result["자본적"])
        result["수익적"] = self.normalizer.to_number_series(result["수익적"])
        
        # 품의금액 합계
        result["품의금액합"] = result["자본적"] + result["수익적"]
        
        # 공장 Forward Fill: 공장명이 병합되어 있을 수 있으므로
        # 빈 공장은 위 행의 공장 사용
        result["공장_산출"] = (
            result["공장"]
            .replace(["", "-"], pd.NA)
            .ffill()
            .fillna("")
        )
        result["공장"] = result["공장_산출"]
        result.drop(columns=["공장_산출"], inplace=True)
        
        # ---------------------------------------------------------
        # 자코드(코드) 산출
        # - 코드 상속(ffill)은 세부행(└ ...)에만 허용
        # - 헤더/구분행(예: ". xxx", 공장 "-" 등)에는 코드가 전파되지 않게 막음
        # ---------------------------------------------------------
        raw_code = result["코드_원본(자코드)"].replace(["", "-"], pd.NA)

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

        # ---------------------------------------------------------
        # 부서코드/유형(계정코드) 산출
        # - 세부행(└)은 부모 행의 부서코드/유형을 상속
        # - 구분행에는 전파하지 않음
        # ---------------------------------------------------------
        raw_dept_code = result["부서코드_원본"].replace(["", "-"], pd.NA)
        raw_acct_code = result["유형_원본(계정코드)"].replace(["", "-"], pd.NA)
        
        ffilled_dept_code = raw_dept_code.ffill()
        ffilled_acct_code = raw_acct_code.ffill()
        
        result["부서코드_산출"] = ""
        result["계정코드_산출"] = ""
        
        # 원본 값이 있는 행
        result.loc[raw_dept_code.notna(), "부서코드_산출"] = raw_dept_code.loc[raw_dept_code.notna()].astype(str)
        result.loc[raw_acct_code.notna(), "계정코드_산출"] = raw_acct_code.loc[raw_acct_code.notna()].astype(str)
        
        # 세부행이면 부모 행에서 상속
        result.loc[mask_inherit, "부서코드_산출"] = ffilled_dept_code.loc[mask_inherit].fillna("").astype(str)
        result.loc[mask_inherit, "계정코드_산출"] = ffilled_acct_code.loc[mask_inherit].fillna("").astype(str)
        
        # 구분행에는 전파 금지
        result.loc[bad_plant & (~is_detail) & raw_dept_code.isna(), "부서코드_산출"] = ""
        result.loc[bad_plant & (~is_detail) & raw_acct_code.isna(), "계정코드_산출"] = ""

        # ---------------------------------------------------------
        # 계정코드 자동 기입: 원본 유형이 비어있으면 금액 기준으로 결정
        #   자본적 > 0 → 955701
        #   수익적 > 0 → 524403
        # ---------------------------------------------------------
        acct_empty = result["계정코드_산출"].isin(["", "-"])
        capex_positive = result["자본적"] > 0
        opex_positive = result["수익적"] > 0

        result.loc[acct_empty & capex_positive, "계정코드_산출"] = "955701"
        result.loc[acct_empty & opex_positive & ~capex_positive, "계정코드_산출"] = "524403"
        # 양쪽 모두 있으면 자본적(955701) 우선 (이미 위에서 설정됨)

        # =========================================================
        # 세부행 합산 + 투자명/부서 forward fill
        # =========================================================

        # 부서 Forward Fill (병합셀/빈칸 대응)
        result["부서_산출"] = (
            result["부서"]
            .replace(["", "-"], pd.NA)
            .ffill()
            .fillna("")
        )
        result["부서"] = result["부서_산출"]

        # 투자명 Forward Fill: 원본 코드 셀이 존재하는 행의 투자명을 기준으로 ffill
        code_cell = result["코드_원본(자코드)"].replace(["", "-"], pd.NA)
        name_base = result["투자명"].where(code_cell.notna(), pd.NA)
        result["투자명"] = name_base.ffill().fillna("")

        # 동일 자코드(+공장) 상세행 합산
        group_cols = ["공장", "부서", "자코드_산출", "부서코드_산출", "계정코드_산출"]
        aggregated = (
            result
            .groupby(group_cols, dropna=False, as_index=False)
            .agg(
                투자명=("투자명", "first"),
                **{
                    "부서코드_원본": ("부서코드_원본", "first"),
                    "유형_원본(계정코드)": ("유형_원본(계정코드)", "first"),
                    "코드_원본(자코드)": ("코드_원본(자코드)", "first"),
                    "자본적": ("자본적", "sum"),
                    "수익적": ("수익적", "sum"),
                    "품의금액합": ("품의금액합", "sum"),
                }
            )
        )

        # 코드_원본(자코드)는 그룹 키(자코드_산출)와 통일
        aggregated["코드_원본(자코드)"] = aggregated["자코드_산출"]

        # 반환 컬럼 정리(이후 단계에서 필요한 컬럼 위주)
        keep_cols = [
            "공장", "부서", "투자명",
            "부서코드_원본", "유형_원본(계정코드)", "코드_원본(자코드)",
            "자본적", "수익적", "품의금액합",
            "자코드_산출", "부서코드_산출", "계정코드_산출",
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
        등록 대상 필터링 (신규 양식)
        
        필터링 기준:
        1. 투자계획("-")이 아닌 실제 투자코드만
        2. 해당 행에 품의금액이 있는 항목만
        3. 자코드_산출 컬럼 사용 (forward fill된 코드)
        4. 합산 행(소계, 공장, 전사 등) 제외
        """
        result = df.copy()
        
        # 합산 행 판별: 부서(팀)가 '전공장' 또는 '전부서'인 행
        def is_summary_row(row) -> bool:
            부서 = str(row.get("부서", "")).strip()
            return 부서 in {"전공장", "전부서"}
        
        # 자코드는 이미 forward fill된 "자코드_산출" 사용
        result["자코드"] = result["자코드_산출"]
        
        # 합산 행 제거
        result["_is_summary"] = result.apply(is_summary_row, axis=1)
        summary_count = result["_is_summary"].sum()
        
        if summary_count > 0:
            logger.info(f"합산 행(소계/합계) 제거: {summary_count}건")
        
        # 필터링: 유효한 자코드가 있고 품의금액이 있고 합산 행이 아닌 것만
        targets = result[
            (result["자코드"] != "") &
            (result["자코드"] != "-") &
            (result["품의금액합"] > 0) &
            (~result["_is_summary"])
        ].copy()
        
        # 임시 컬럼 제거
        targets.drop(columns=["_is_summary"], inplace=True, errors='ignore')
        
        # 코드 분류 플래그 추가
        targets["코드번호"] = targets["자코드"].apply(
            ActualFileHandler._get_code_number
        )
        
        targets["코드없이금액만"] = (
            (targets["코드_원본(자코드)"] == "") | 
            (targets["코드_원본(자코드)"] == "-")
        ) & (targets["자코드"] != "")
        
        # 통계 정보
        total_targets = len(targets)
        no_code_but_amount = targets["코드없이금액만"].sum()
        
        logger.info(
            f"등록 대상 필터링 완료: "
            f"총 {len(df)}행 중 {total_targets}행 추출 "
            f"(코드없이금액: {no_code_but_amount}건)"
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
        self.mis_handler = MISFileHandler(config, self.normalizer)
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
        # MIS 데이터 로드
        mis_path = self.mis_handler.pick_latest_file()
        mis_df = self.mis_handler.load_file(mis_path)

        # 실제 데이터 로드
        targets, actual_all = self.actual_handler.load_file(
            Path(actual_path), 
            actual_sheet_name
        )

        # =========================================================
        # MIS 입력오류 최종 판정: "본사 단독 + 금액 입력" 후보 중
        # 실제 투자이력(등록대상)에서 '생산 공장'으로 등장하는 코드만 오류로 확정
        # =========================================================
        cand_codes = set(mis_df.attrs.get("hq_amount_error_candidate_codes", []))
        final_bad_codes = set()
        if cand_codes:
            prod_mask = (
                (targets["공장"] != "본사") &
                (targets["공장"] != "-") &
                (targets["공장"].astype(str).str.strip() != "") &
                (~targets["공장"].astype(str).str.contains("연구소", na=False))
            )
            prod_codes = set(targets.loc[prod_mask, "자코드"].astype(str))
            final_bad_codes = cand_codes & prod_codes

        if final_bad_codes:
            before = len(mis_df)
            mis_df = mis_df[~mis_df["자코드"].isin(final_bad_codes)].copy()
            removed = before - len(mis_df)
            logger.warning(f"MIS 입력오류 확정(본사 단독+금액입력): 코드 {len(final_bad_codes)}개, 행 {removed}건 제외 → 전산미등록으로 처리")
            mis_df.attrs["hq_amount_error_df"] = mis_df.attrs.get("hq_amount_error_candidate_df", pd.DataFrame()).copy()
            mis_df.attrs["hq_amount_error_codes"] = sorted(final_bad_codes)
        else:
            mis_df.attrs["hq_amount_error_df"] = pd.DataFrame(columns=mis_df.columns)
            mis_df.attrs["hq_amount_error_codes"] = []

        # 집계는 최종 MIS(필터 후) 기준으로 수행
        mis_aggregated = self.mis_handler.aggregate_by_key(mis_df)

        # =========================================================
        # 병합 로직:
        # 1) 자코드 + 공장으로 정확 매칭
        # 2) 공장 불일치 시 자코드만으로 재매칭 (전사 코드 등)
        # =========================================================
        
        # Step 1: 자코드 + 공장으로 병합
        merged = targets.merge(
            mis_aggregated,
            on=["자코드", "공장"],
            how="left",
            suffixes=('', '_MIS')
        )
        
        # Step 2: 매칭 실패한 경우, 자코드만으로 재시도
        unmatched = merged[merged["MIS행수"].isna()].copy()
        
        if len(unmatched) > 0:
            logger.info(
                f"공장 불일치 {len(unmatched)}건 → 자코드만으로 재매칭 시도"
            )
            
            retry_cols = ["자코드", "MIS행수", "계정코드입력여부", "계정코드값", "투자계획금액", "품의승인금액"]
            remerged = unmatched[["자코드"]].merge(
                mis_aggregated[retry_cols].drop_duplicates(subset=["자코드"]),
                on=["자코드"],
                how="left",
                suffixes=('', '_retry')
            )

            for idx in unmatched.index:
                자코드 = unmatched.loc[idx, "자코드"]
                match_result = remerged[remerged["자코드"] == 자코드]

                if len(match_result) > 0 and not pd.isna(match_result.iloc[0]["MIS행수"]):
                    row0 = match_result.iloc[0]
                    merged.loc[idx, "MIS행수"] = row0["MIS행수"]
                    merged.loc[idx, "계정코드입력여부"] = row0["계정코드입력여부"]
                    merged.loc[idx, "계정코드값"] = row0["계정코드값"]
                    merged.loc[idx, "투자계획금액"] = row0["투자계획금액"]
                    merged.loc[idx, "품의승인금액"] = row0["품의승인금액"]
            
            rematched = merged.loc[unmatched.index, "MIS행수"].notna().sum()
            if rematched > 0:
                logger.info(f"재매칭 성공: {rematched}건")
        
        # 검증 1: 전산 미등록
        missing_in_mis = merged[merged["MIS행수"].isna()].copy()

        # 검증 2: 계정코드 미입력
        account_missing = merged[
            (merged["MIS행수"].notna()) &
            (merged["계정코드입력여부"] == False)
        ].copy()

        # 검증 3: 품의승인금액 미입력
        # MIS에 등록은 됐지만 품의승인금액이 0이고 실제 파일에는 품의금액이 있는 경우
        # (실제 파일=천원 단위, MIS=원 단위)
        approval_missing = merged[
            (merged["MIS행수"].notna()) &
            (merged["품의승인금액"].fillna(0) == 0) &
            (merged["품의금액합"] > 0)
        ].copy()
        if not approval_missing.empty:
            approval_missing["MIS품의승인금액(천원)"] = approval_missing["품의승인금액"].fillna(0) / 1000
            approval_missing["차액(천원)"] = (
                approval_missing["품의금액합"] - approval_missing["MIS품의승인금액(천원)"]
            )

        # 검증 4: 계정코드 오입력
        # 자본적 > 0 (수익적 0) → 955701
        # 수익적 > 0 (자본적 0) → 524403
        # 둘 다 > 0 → 955701, 524403 모두 필요
        def expected_acct(row) -> set:
            capex = float(row.get("자본적", 0) or 0)
            opex = float(row.get("수익적", 0) or 0)
            expected = set()
            if capex > 0:
                expected.add("955701")
            if opex > 0:
                expected.add("524403")
            return expected

        def acct_mismatch(row) -> bool:
            if pd.isna(row.get("MIS행수")):
                return False
            expected = expected_acct(row)
            if not expected:
                return False
            actual_str = str(row.get("계정코드값", "") or "").strip()
            if not actual_str:
                # 계정코드 자체가 비어있는 경우는 '계정코드미입력'에서 처리
                return False
            actual_set = {c.strip() for c in actual_str.split(",") if c.strip()}
            # 기대 계정코드가 MIS 계정코드 집합에 모두 포함되어야 정상
            return not expected.issubset(actual_set)

        merged["기대계정코드"] = merged.apply(
            lambda r: ",".join(sorted(expected_acct(r))), axis=1
        )
        merged["계정코드오입력"] = merged.apply(acct_mismatch, axis=1)
        account_mismatch = merged[merged["계정코드오입력"]].copy()

        # 검증 5: MIS 내부 - 투자계획금액 vs 품의승인금액 불일치
        # (실제 파일과 무관, MIS 자체 검증)
        plan_approval_mismatch = self._detect_plan_approval_mismatch(mis_df)

        # 검증 6: MIS 내부 - 151+ 자코드의 모코드 미입력
        # 자코드가 XX151 이상인 경우 모코드가 입력되어 있어야 함
        missing_parent_code = self._detect_missing_parent_code(mis_df)

        # 검증 7: MIS 내부 - 투자자(담당자) 미입력
        missing_investor = self._detect_missing_name_column(mis_df, "투자자")

        logger.info(
            f"검증 완료 - 전산미등록: {len(missing_in_mis)}건, "
            f"계정코드미입력: {len(account_missing)}건, "
            f"품의승인금액미입력: {len(approval_missing)}건, "
            f"계정코드오입력: {len(account_mismatch)}건, "
            f"계획-승인불일치: {len(plan_approval_mismatch)}건, "
            f"모코드미입력(151+): {len(missing_parent_code)}건, "
            f"투자자미입력: {len(missing_investor)}건"
        )

        return {
            "mis_path": mis_path,
            "actual_path": Path(actual_path),
            "mis_data": mis_df,
            "mis_aggregated": mis_aggregated,
            "targets": targets,
            "actual_all": actual_all,
            "merged": merged,
            "missing_in_mis": missing_in_mis,
            "account_missing": account_missing,
            "approval_missing": approval_missing,
            "account_mismatch": account_mismatch,
            "plan_approval_mismatch": plan_approval_mismatch,
            "missing_parent_code": missing_parent_code,
            "missing_investor": missing_investor,
        }

    def _detect_plan_approval_mismatch(self, mis_df: pd.DataFrame) -> pd.DataFrame:
        """
        MIS 내부 검증: 투자계획금액 vs 품의승인금액 불일치
        - 양쪽 모두 > 0이면서 금액이 다른 행을 추출
        - 한쪽이 0인 경우는 '품의승인금액 미입력' 등 다른 검증에서 처리
        """
        if mis_df is None or mis_df.empty:
            return pd.DataFrame()

        needed = ["투자계획금액", "품의승인금액"]
        if not all(c in mis_df.columns for c in needed):
            return pd.DataFrame()

        work = mis_df.copy()
        work["_계획"] = self.normalizer.to_number_series(work["투자계획금액"])
        work["_승인"] = self.normalizer.to_number_series(work["품의승인금액"])

        mask = (work["_계획"] > 0) & (work["_승인"] > 0) & (work["_계획"] != work["_승인"])
        out = work[mask].copy()

        if out.empty:
            return out

        out["계획-승인차액"] = out["_계획"] - out["_승인"]
        # 보조 컬럼 제거(정렬용 차액은 남김)
        out.drop(columns=["_계획", "_승인"], inplace=True)

        # 자코드 기준 정렬
        sort_cols = [c for c in ["자코드", "공장"] if c in out.columns]
        if sort_cols:
            out = out.sort_values(sort_cols).reset_index(drop=True)
        return out

    def _detect_missing_parent_code(self, mis_df: pd.DataFrame) -> pd.DataFrame:
        """
        MIS 내부 검증: 151+ 자코드인데 모코드가 비어있는 경우
        - 자코드가 [A-Z]{2}\\d{3} 형식이면서 끝 3자리 숫자가 151 이상
        - 모코드 컬럼이 빈값/하이픈인 행을 추출
        """
        if mis_df is None or mis_df.empty or "모코드" not in mis_df.columns:
            return pd.DataFrame()

        work = mis_df.copy()

        def extract_num(code) -> int:
            m = re.match(r'^[A-Za-z]{2}(\d{3})$', str(code).strip())
            return int(m.group(1)) if m else -1

        work["_자코드숫자"] = work["자코드"].apply(extract_num)
        work["_모코드비움"] = (
            work["모코드"].fillna("").astype(str).str.strip().isin(["", "-", "nan", "NaN"])
        )

        mask = (work["_자코드숫자"] >= 151) & (work["_모코드비움"])
        out = work[mask].copy()
        out.drop(columns=["_자코드숫자", "_모코드비움"], inplace=True)

        sort_cols = [c for c in ["자코드", "공장"] if c in out.columns]
        if sort_cols:
            out = out.sort_values(sort_cols).reset_index(drop=True)
        return out

    def _detect_missing_name_column(self, mis_df: pd.DataFrame, col_name: str) -> pd.DataFrame:
        """
        MIS 내부 검증: 지정한 컬럼이 입력되지 않은 행 추출
        - 투자명, 투자자(담당자) 등 이름 계열 컬럼에 사용
        """
        if mis_df is None or mis_df.empty or col_name not in mis_df.columns:
            return pd.DataFrame()

        work = mis_df.copy()
        empty_mask = (
            work[col_name].fillna("").astype(str).str.strip().isin(["", "-", "nan", "NaN"])
        )
        out = work[empty_mask].copy()

        sort_cols = [c for c in ["자코드", "공장"] if c in out.columns]
        if sort_cols:
            out = out.sort_values(sort_cols).reset_index(drop=True)
        return out
    
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
            "종합정보시스템 파일": validation_result["mis_path"].name,
            "실제 투자실적 파일": validation_result["actual_path"].name,
            "실제 등록대상 행수": int(len(targets)),
            "실제 등록대상 자코드수": int(targets["자코드"].nunique()),
            "종합정보 자코드수": int(
                validation_result["mis_data"]["자코드"].nunique()
            ),
            "전산미등록": int(len(validation_result["missing_in_mis"])),
            "계정코드미입력": int(len(validation_result["account_missing"])),
            "품의승인금액미입력": int(len(validation_result.get("approval_missing", pd.DataFrame()))),
            "계정코드오입력": int(len(validation_result.get("account_mismatch", pd.DataFrame()))),
            "계획-승인금액불일치": int(len(validation_result.get("plan_approval_mismatch", pd.DataFrame()))),
            "모코드미입력(151+)": int(len(validation_result.get("missing_parent_code", pd.DataFrame()))),
            "투자자미입력": int(len(validation_result.get("missing_investor", pd.DataFrame()))),
            "실제파일_사용시트": targets.attrs.get("used_sheet", ""),
            "실제파일_헤더행(0기준)": targets.attrs.get("header_row", ""),
        }])

        # Excel 파일 작성
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary.to_excel(writer, index=False, sheet_name="요약")

            validation_result["missing_in_mis"].to_excel(
                writer, index=False, sheet_name="전산미등록"
            )

            validation_result["account_missing"].to_excel(
                writer, index=False, sheet_name="계정코드미입력"
            )

            approval_missing = validation_result.get("approval_missing", pd.DataFrame())
            if not approval_missing.empty:
                approval_missing.to_excel(
                    writer, index=False, sheet_name="품의승인금액미입력"
                )

            account_mismatch = validation_result.get("account_mismatch", pd.DataFrame())
            if not account_mismatch.empty:
                account_mismatch.to_excel(
                    writer, index=False, sheet_name="계정코드오입력"
                )

            plan_approval_mismatch = validation_result.get("plan_approval_mismatch", pd.DataFrame())
            if not plan_approval_mismatch.empty:
                plan_approval_mismatch.to_excel(
                    writer, index=False, sheet_name="계획-승인금액불일치"
                )

            missing_parent_code = validation_result.get("missing_parent_code", pd.DataFrame())
            if not missing_parent_code.empty:
                missing_parent_code.to_excel(
                    writer, index=False, sheet_name="모코드미입력(151+)"
                )

            missing_investor = validation_result.get("missing_investor", pd.DataFrame())
            if not missing_investor.empty:
                missing_investor.to_excel(
                    writer, index=False, sheet_name="투자자미입력"
                )

            validation_result["mis_aggregated"].to_excel(
                writer, index=False, sheet_name="종합정보 집계(자코드+공장)"
            )

            # MIS 입력오류(본사 단독 + 금액 입력) 목록
            mis_data = validation_result.get("mis_data")
            hq_err_df = None
            if mis_data is not None:
                hq_err_df = mis_data.attrs.get("hq_amount_error_df")
            if isinstance(hq_err_df, pd.DataFrame) and not hq_err_df.empty:
                hq_err_df.to_excel(writer, index=False, sheet_name="MIS입력오류(본사금액)")

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
    print(f"전산미등록: {len(result['missing_in_mis'])}건")
    print(f"계정코드미입력: {len(result['account_missing'])}건")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    # 기본 동작:
    # - GUI 환경이면 파일 선택창 사용 가능
    # - 헤드리스/서버 환경에서는 actual_path를 직접 지정해서 실행하세요.
    #
    # 예)
    #   python investment_validator.py --actual "/path/to/actual.xlsx" --MIS "/path/to/MIS.xlsx"
    import argparse

    parser = argparse.ArgumentParser(description="투자실적 검증 자동화")
    parser.add_argument("--actual", dest="actual_path", required=False, help="실제 투자실적 파일 경로")
    parser.add_argument("--sheet", dest="actual_sheet_name", default=0, help="시트 이름 또는 인덱스 (기본 0)")
    parser.add_argument("--MIS", dest="mis_path", required=False, help="MIS(종합정보) 파일 경로 (미지정 시 최신 파일 자동 선택)")
    args = parser.parse_args()

    # Config 경로는 기본값 유지. 단, --MIS를 주면 해당 파일을 우선 사용
    cfg = Config()
    validator = InvestmentValidator(cfg)

    if args.actual_path is None:
        # GUI 환경이면 파일 선택 다이얼로그
        main(actual_path=None, actual_sheet_name=args.actual_sheet_name)
    else:
        # MIS 파일 강제 지정 옵션 지원
        if args.mis_path:
            MIS_p = Path(args.mis_path)
            mis_df = validator.mis_handler.load_file(MIS_p)
            MIS_agg = validator.mis_handler.aggregate_by_key(mis_df)
            targets, actual_all = validator.actual_handler.load_file(Path(args.actual_path), args.actual_sheet_name)

            merged = targets.merge(MIS_agg, on=["자코드", "공장"], how="left", suffixes=('', '_MIS'))
            missing_in_mis = merged[merged["MIS행수"].isna()].copy()
            account_missing = merged[(merged["MIS행수"].notna()) & (merged["계정코드입력여부"] == False)].copy()

            result = {
                "mis_path": MIS_p,
                "actual_path": Path(args.actual_path),
                "mis_data": mis_df,
                "mis_aggregated": MIS_agg,
                "targets": targets,
                "actual_all": actual_all,
                "merged": merged,
                "missing_in_mis": missing_in_mis,
                "account_missing": account_missing,
            }
            out = validator.generate_report(result)
            print(out)
        else:
            main(actual_path=Path(args.actual_path), actual_sheet_name=args.actual_sheet_name)
