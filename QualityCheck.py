import os, re, json, copy, time, datetime, traceback
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
        self.logger = Logger(proc_name='Quality Check', log_folder_path = self.PATH['LOG'])
        
    def data_check(self):
        # 데이터 존재 여부 확인
        try:
            # 데이터 파일 리스트 확인 (숨김파일 제거)
            files = [file for file in os.listdir(self.PATH['DATA']) if not file.startswith('.')]
            
            # 데이터 파일이 존재하지 않을 경우 에러 로그 기록
            if len(files)==0:
                self.logger.error(f"QC를 수행할 데이터 파일이 존재하지 않습니다.")
                return
                
            self.logger.info(f"총 {len(files):,} 개의 데이터 파일이 존재합니다.")            
            # for file in files:
            #     self.logger.info(f"{file}")
            
        except Exception as e:
            self.logger.error(e)
            err_msg = traceback.format_exc()
            self.logger.error(err_msg)            
            
        
    def document_check(self):
        # 정의서 파일 존재 여부 확인

        doc_list = ['테이블정의서', '컬럼정의서', '코드정의서']        
        self.documents={}
        
        for doc in doc_list:
            self.documents[doc]=[file for file in os.listdir(self.PATH['DOCS']) if doc in file]
            if len(self.documents[doc])
            
            
    def dtype_check(self):
        # 데이터 타입 Custom 기능
        
        return
        
    def na_check(self):
        # 결측값 Custom 기능
        
        return
        
    def run(self):
        # QC 수행
        
        return
        
    def save(self):
        # 결과 저장
        
        return        
        
# if __name__=='__main__':
    
#     Process=QualityCheck()
#     Process.data_check()