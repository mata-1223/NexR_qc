# NexR_qc
[![PyPI version](https://badge.fury.io/py/NexR-qc.svg)](https://badge.fury.io/py/NexR-qc)
<br><br>

## 요구사항
- python >= 3.6
- numpy
- pandas
- openpyxl
<br>

## 설치

### pip 설치
```
#!/bin/bash
pip install NexR_qc
```

### 디렉토리 기본 구성
- documents 하위 항목(테이블정의서, 컬럼정의서, 코드정의서)은 필수 항목은 아니지만, 테이블별 정확한 정보를 얻기위해서 작성되는 문서임 ([Github 링크](https://github.com/mata-1223/NexR_qc)의 document 폴더 내 문서 양식 참고)
- log, output 폴더는 초기에 생성되어 있지않아도 수행 결과로 자동 생성됨
- config.json 파일은 데이터 내 결측값을 커스텀하기 위한 파일로 초기에 생성되어 있지않아도 수행 결과로 자동 생성됨 (결측처리 default 값: "?", "na", "null", "Null", "NULL", " ", "[NULL]")

```
.
├── data/ (optional)
│   ├── 데이터_001.csv
│   ├── 데이터_002.csv
│   ├── 데이터_003.xlsx
│   ├── ...
├── documents/
│   ├── 테이블정의서.xlsx
│   ├── 컬럼정의서.xlsx
│   └── 코드정의서.xlsx
├── log/
│   ├── QualityCheck_yyyymmdd_hhmmss.log
│   ├── ...
├── output/
│   └── QC결과서_yyyymmdd_hhmmss.xlsx
└── config.json
``` 
<br>

## 예제 실행 
```
#!bin/usr/python3
from NexR_qc.QualityCheck import *

# 데이터 불러오기 (데이터 파일 활용 시)
PathDict = {}
PathDict["ROOT"] = os.getcwd()
PathDict["DATA"] = os.path.join(PathDict["ROOT"], "data")  # 데이터 파일이 있는 디렉토리 경로

# 데이터 불러오기 (DB 활용시)
# DB에 적재된 데이터를 데이터프레임 형태로 불러와 하단 DataDict 형태에 맞게 준비

DataDict = {}  # DataDict: 데이터명(key)-데이터프레임(value)로 이루어짐
for path in [i for i in os.listdir(PathDict["DATA"]) if not i.startswith(".")]:
    data_name = os.path.splitext(os.path.basename(path))[0].upper()
    DataDict[data_name] = pd.read_csv(os.path.join(PathDict["DATA"], path))

Process = QualityCheck(DataDict)
Process.data_check()
Process.document_check()
Process.na_check()
Process.run()
Process.save()
```

<br>

## Input / Output 정보

### Input
* 데이터 타입: Dictionary 형태
	* 상세 형상: {data_name1: Dataframe1, data_name2: Dataframe2,…}
		* data_name: 데이터 테이블명 or 데이터 파일명 
		* Dataframe: 데이터를 불러온 Dataframe 형상
* 예시
![NexR_qc_Info_002](https://github.com/mata-1223/NexR_qc/assets/131343466/5e28e8bf-37f2-4cc0-acca-c288bfbd5ccb)

### Output
* 결과 파일 경로: output/QC_결과서.xlsx
* 예시
1) 예시 1: 테이블 리스트 시트
![NexR_qc_Info_003](https://github.com/mata-1223/NexR_qc/assets/131343466/54605ebe-d45c-4ba9-b219-dd177e08a6b7)

2) 예시 2: 데이터 별 QC 수행 결과 시트
![NexR_qc_Info_001](https://github.com/mata-1223/NexR_qc/assets/131343466/a1613944-4812-40a2-9ec3-6452c104a96b)