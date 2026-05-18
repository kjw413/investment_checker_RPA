# 빠른 시작 가이드

## 🚀 10초 안에 시작하기

### 1. 기본 실행 (가장 간단)
```python
from investment_validator_refactored import main

# GUI 파일 선택 다이얼로그가 나타납니다
main()
```

---

## 📋 사전 준비

### 필수 패키지 설치
```bash
pip install pandas openpyxl
```

### 선택 패키지 (테스트용)
```bash
pip install pytest pytest-cov
```

---

## 💻 사용 예시

### 예시 1: GUI로 파일 선택
```python
from pathlib import Path
from investment_validator_refactored import main

# 파일 선택 다이얼로그의 초기 디렉토리 지정
main(initial_dir=Path(r"E:\03. 투자관리\투자 계획_실적 조회\투자품의 현항"))
```

### 예시 2: 파일 경로 직접 지정
```python
from investment_validator_refactored import main

# 파일 경로를 직접 지정
main(
    actual_path=r"E:\투자관리\실제투자실적_2024.xlsx",
    actual_sheet_name=0  # 첫 번째 시트
)
```

### 예시 3: 시트 이름 지정
```python
from investment_validator_refactored import main

main(
    actual_path="실제투자실적_2024.xlsx",
    actual_sheet_name="투자현황"  # 시트 이름으로 지정
)
```

### 예시 4: 커스텀 설정
```python
from pathlib import Path
from investment_validator_refactored import Config, InvestmentValidator

# 설정 커스터마이징
config = Config(
    erp_dir=Path("./my_erp_files"),        # ERP 파일 디렉토리
    output_dir=Path("./my_results"),        # 결과 저장 디렉토리
    max_scan_rows=300,                      # 헤더 탐색 최대 행 수
    min_header_hits=4,                      # 헤더 인식 최소 키워드 수
)

# 검증 실행
validator = InvestmentValidator(config)
result = validator.validate(
    actual_path="실제투자실적.xlsx",
    actual_sheet_name=0
)

# 결과 확인
print(f"전산 미등록: {len(result['missing_in_erp'])}건")
print(f"계정코드 미입력: {len(result['account_missing'])}건")

# 리포트 생성
output_path = validator.generate_report(result)
print(f"결과 파일: {output_path}")
```

### 예시 5: 프로그래밍 방식으로 활용
```python
from investment_validator_refactored import InvestmentValidator, Config

# 검증 실행
validator = InvestmentValidator(Config())
result = validator.validate("actual.xlsx")

# 결과 데이터 활용
missing = result['missing_in_erp']
account_missing = result['account_missing']

# 커스텀 처리
for idx, row in missing.iterrows():
    print(f"미등록: {row['투자명']} ({row['자코드']})")

for idx, row in account_missing.iterrows():
    print(f"계정코드 누락: {row['투자명']} ({row['자코드']})")

# 리포트 생성 (선택)
validator.generate_report(result)
```

---

## 📂 디렉토리 구조 예시

```
프로젝트/
├── investment_validator_refactored.py  # 메인 스크립트
├── test_investment_validator.py        # 테스트 (선택)
├── 종합정보 투자실적/                   # ERP 파일 폴더
│   ├── ERP_2024_01.xlsx
│   ├── ERP_2024_02.xlsx
│   └── ERP_2024_03.xlsx               # 가장 최근 파일 자동 선택
└── result/                             # 결과 폴더 (자동 생성)
    └── 투자실적_점검결과_20241222_153045.xlsx
```

---

## 🎯 출력 결과

### 콘솔 출력
```
INFO - ERP 파일 선택: ERP_2024_03.xlsx
INFO - ERP 파일 로딩 중: C:\...\ERP_2024_03.xlsx
INFO - ERP 데이터 로딩 완료: 524행
INFO - ERP 집계 완료: 312개 키
INFO - 실제 파일 로딩 중: C:\...\실제투자실적_2024.xlsx
INFO - 헤더 행 발견: 2행 (매칭: {'공장', '부서', '투자명', '투자코드', '자본적', '수익적'})
INFO - 시트 로딩 성공: 'Sheet1' (헤더: 2행)
INFO - 실제 데이터 로딩 완료: 전체 287행, 등록대상 256행
INFO - 검증 완료 - 전산미등록: 15건, 계정코드미입력: 23건
INFO - 리포트 생성 완료: C:\...\result\투자실적_점검결과_20241222_153045.xlsx

============================================================
투자실적 검증 완료
============================================================
결과 파일: C:\...\result\투자실적_점검결과_20241222_153045.xlsx
전산미등록: 15건
계정코드미입력: 23건
============================================================
```

### Excel 결과 파일 구조
```
Sheet 1: 요약
- 종합정보시스템 파일명
- 실제 투자실적 파일명
- 등록 대상 행수/자코드수
- 전산미등록 건수
- 계정코드미입력 건수

Sheet 2: 전산미등록
- 전산에 등록되지 않은 투자 항목 목록

Sheet 3: 계정코드미입력
- 전산에는 등록되었으나 계정코드가 누락된 항목

Sheet 4: 종합정보 집계(자코드+공장)
- ERP 데이터 집계 결과

Sheet 5: 실제대상목록(필터후)
- 검증 대상으로 필터링된 실제 투자 항목

Sheet 6: 병합결과(검토용)
- 실제 데이터와 ERP 데이터 병합 결과
```

---

## ⚙️ 설정 옵션

### Config 클래스 주요 파라미터

```python
@dataclass
class Config:
    # 디렉토리 설정
    base_dir: Path              # 기본 디렉토리 (default: 스크립트 상위 폴더)
    erp_dir: Path               # ERP 파일 폴더
    output_dir: Path            # 결과 저장 폴더
    
    # 헤더 탐지 설정
    max_scan_rows: int = 200    # 헤더 탐색할 최대 행 수
    min_header_hits: int = 3    # 헤더로 인식할 최소 키워드 매칭 수
    
    # 정규화 설정
    null_tokens: frozenset[str] = frozenset({"nan", "none", "null"})
    
    # 컬럼 별칭 (헤더 탐지용)
    column_aliases: dict[str, list[str]]  # 커스터마이징 가능
```

---

## 🐛 문제 해결

### Q1: "tkinter를 사용할 수 없는 환경입니다" 오류
```python
# 해결: 파일 경로를 직접 지정
main(actual_path="path/to/file.xlsx")
```

### Q2: "헤더 행을 찾지 못했습니다" 오류
```python
# 해결 1: max_scan_rows 증가
config = Config(max_scan_rows=500)

# 해결 2: min_header_hits 감소
config = Config(min_header_hits=2)

# 해결 3: 시트 이름 직접 지정
main(actual_path="file.xlsx", actual_sheet_name="특정시트명")
```

### Q3: "ERP 파일을 찾을 수 없습니다" 오류
```python
# 해결: ERP 파일 경로 지정
config = Config(erp_dir=Path("./custom_erp_dir"))
validator = InvestmentValidator(config)
```

### Q4: 특정 컬럼명이 달라서 인식 실패
```python
# 해결: column_aliases 커스터마이징
config = Config()
config.column_aliases["공장"].append("사업소")  # "사업소"도 공장으로 인식
config.column_aliases["투자명"].append("프로젝트명")
```

---

## 🧪 테스트 실행

### 기본 테스트
```bash
pytest test_investment_validator.py -v
```

### 커버리지 포함
```bash
pytest test_investment_validator.py --cov=investment_validator_refactored --cov-report=html
```

### 특정 테스트만 실행
```bash
# DataNormalizer 테스트만
pytest test_investment_validator.py::TestDataNormalizer -v

# 특정 테스트 케이스만
pytest test_investment_validator.py::TestDataNormalizer::test_normalize_code_series_removes_spaces -v
```

---

## 📞 추가 지원

### 상세 문서
- **README.md**: 프로젝트 개요 및 리팩터링 설명
- **REFACTORING_COMPARISON.md**: 기존 vs 리팩터링 코드 비교
- **SUMMARY.md**: 전체 프로젝트 요약

### 문의
- 버그 발견 시: GitHub Issues
- 기능 제안: Pull Request
- 긴급 문의: 담당자 이메일

---

## 🎓 다음 단계

1. ✅ 이 가이드대로 기본 실행 테스트
2. 📚 README.md에서 리팩터링 내용 확인
3. 🔍 REFACTORING_COMPARISON.md에서 코드 비교
4. 🧪 test_investment_validator.py에서 테스트 예시 확인
5. 🚀 실제 데이터로 검증 수행

---

**즐거운 코딩 되세요! 🎉**
