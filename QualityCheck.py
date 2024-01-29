import os, re, json, copy, time, datetime, traceback
import unicodedata
from pathlib import Path
import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook
from utils.Logging import *
from utils.Timer import *

class QualityCheck:
    
    def __init__(self):
        # 초기 디렉토리 세팅
        self.PATH={}
        self.PATH['ROOT']=os.getcwd()
        self.PATH['DATA']=os.path.join(self.PATH['ROOT'], 'data')
        self.PATH['DOCS']=os.path.join(self.PATH['ROOT'], 'documents')
        self.PATH['LOG']=os.path.join(self.PATH['ROOT'], 'log')
        self.PATH['OUTPUT']=os.path.join(self.PATH['ROOT'], 'output')
        
        for folder in self.PATH.keys():
            # 구성 폴더가 없을 경우, 생성
            if not os.path.exists(self.PATH[folder]):
                Path(self.PATH[folder]).mkdir(parents=True, exist_ok=True)
                
        # 로그 구성
        self.logger=Logger(proc_name='Quality Check', log_folder_path = self.PATH['LOG'])
        
        # Config 파일 불러오기
        with open(os.path.join(self.PATH['ROOT'], 'config.json'), 'r') as f:
            self.config=json.load(f)
        
    def data_check(self):
        # 데이터 존재 여부 확인
        # try:
            # 데이터 파일 리스트 확인 (숨김파일 제거)
        files=[file for file in os.listdir(self.PATH['DATA']) if not file.startswith('.')]
        
        # 데이터 파일이 존재하지 않을 경우 에러 로그 기록
        if len(files)==0:
            self.logger.error(f"QC를 수행할 데이터 파일이 존재하지 않습니다.")
            return
            
        self.logger.info(f"총 {len(files):,} 개의 데이터 파일이 존재합니다.")
        # except Exception as e:
        #     self.logger.error(e)
        #     err_msg=traceback.format_exc()
        #     self.logger.error(err_msg)            
            
    def document_check(self):
        # 정의서 파일 존재 여부 확인

        doc_list = ['테이블정의서', '컬럼정의서', '코드정의서']        
        self.documents={}
        
        for doc in doc_list:
            self.documents[doc]=[file for file in os.listdir(self.PATH['DOCS']) if unicodedata.normalize('NFC', doc) in unicodedata.normalize('NFC', file)]
        
        if len(self.documents.keys())==0:
            self.logger.info(f"참고할 문서 파일이 없습니다.")
        else:
            self.logger.info(f"참고할 문서 파일이 {len(self.documents.keys())} 개 있습니다.")
            for k, v in self.documents.items():
                self.logger.info(f"{k} 참고 문서 파일: {os.path.join(self.PATH['DOCS'], v[0])}")
            
    def dtype_check(self):
        # 데이터 타입 Custom 기능
        
        return
        
    def na_check(self):
        # 결측값 Custom 기능
        self.naList = self.config['naList']
        self.logger.info(f"현재 결측값으로 등록된 값은 다음과 같습니다.")
        self.logger.info(f"{self.na_list}")
        self.logger.info(f"결측값 추가 등록을 원하시면 config.json 파일 내 na_list 값에 추가 시 반영됩니다.")
        
    def run(self):
        # QC 수행
        
        self.logger.info(f"데이터 QC를 수행합니다.")
        
        # Step 1: 데이터 파일 불러오기
        data_path=os.listdir(self.PATH['DATA'])[0]
        ext=os.path.splitext(data)
        
        if ext=='csv':
            data=pd.read_csv()
        
        
        return
        
    def save(self):
        # 결과 저장
        
        return        