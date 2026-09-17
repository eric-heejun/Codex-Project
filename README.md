# Codex-Project

## 발주서 분류기

최신 발주서 분류기 소스는 아래 폴더에 있습니다.

`release/발주서분류머신_오프라인_20260619/app`

주요 파일:

- `src/classify_orders.py`: 공급처 분류
- `src/export_supplier_orders.py`: 공급처별 발주서 생성
- `src/train_supplier_rules.py`: 기존 분류 결과 검증
- `config/supplier_rules.json`: 상품 및 공급처 규칙
- `templates`: 공급처별 엑셀 양식

현재 주요 공급처 규칙에는 다음 내용이 반영되어 있습니다.

- 팜하우스, 팜마우스, 모니카: 메르시데코
- 카프리섬머, 카프리 썸머, 카프리 시어서커: 혜경산업
- 29CM 주문 옵션의 컬러 및 사이즈 형식 인식
- 룬 울트라 퍼포먼스의 사각형, 원형, 타원형, 러너, 반원매트 인식

소스 체크아웃에서 실행하려면 Python 3.10 이상 환경에서 의존성을 설치합니다.

```powershell
py -m pip install -r "release/발주서분류머신_오프라인_20260619/app/requirements.txt"
py "release/발주서분류머신_오프라인_20260619/app/src/export_supplier_orders.py" --input "주문서.xlsx"
```

결과는 `app/data/output/발주서/MMDD` 아래에 생성됩니다. 내장 Python이 포함된 일반 PC용 배포 ZIP은 GitHub에 올리지 않고 별도로 관리합니다.
