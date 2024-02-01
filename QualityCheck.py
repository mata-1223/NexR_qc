import os, re, json, copy, time, datetime, traceback
import unicodedata
from pathlib import Path
import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.worksheet.dimensions import ColumnDimension
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
        self.logger_save=False # 로그 파일 생성 여부 (True: 로그 파일 생성 / False: 로그 파일 미생성)
        self.logger=Logger(proc_name='Quality Check', log_folder_path = self.PATH['LOG'], save=self.logger_save)
        
        # Config 파일 불러오기
        with open(os.path.join(self.PATH['ROOT'], 'config.json'), 'r') as f:
            self.config=json.load(f)
        
        self.readFunc={}
        self.readFunc['.csv']=pd.read_csv
        self.readFunc['.xlsx']=pd.read_excel
        
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
        self.logger.info(f"{self.naList}")
        self.logger.info(f"결측값 추가 등록을 원하시면 config.json 파일 내 naList 값에 추가 시 반영됩니다.")
        
    def run(self):
        
        # QC 수행
        self.logger.info(f"데이터 QC를 수행합니다.")
        
        # Step 1: 데이터 파일 불러오기
        data_path=[file for file in os.listdir(self.PATH['DATA']) if not file.startswith('.')][0]
        ext=os.path.splitext(data_path)[-1]
        
        data=self.readFunc[ext](os.path.join(self.PATH['DATA'], data_path), na_values=self.naList)
        
        self.logger.info('='*50)
        self.logger.info(f"{data_path} 파일 불러오기 성공")
        self.logger.info(f"데이터 Shape:{data.shape}")
        
        # Step 2: 결과 항목 값 세팅        
        self.RelCategory={'공통': ['No', '컬럼 영문명', '컬럼 한글명', '데이터 타입', 'null 개수', '%null', '적재건수', '%적재건수'],
                    '연속형': ['최솟값', '최댓값', '평균', '표준편차', '중위수'],
                    '범주형': ['범주수', '범주', '%범주', '정의된 범주 외', '정의된 범주 외 수', '최빈값', '최빈값 수', '%최빈값'],
                    '비고': ['비고']}
            
        self.ResultDict={f'{idx:03d}': {'공통': {'No': f'{idx:03d}', '컬럼 영문명': col}} for idx, col in enumerate(data.columns)}
        
        # Step 3-1: 공통 영역 QC 수행
        for idx in self.ResultDict.keys():
            col=self.ResultDict[idx]['공통']['컬럼 영문명']
        
            self.ResultDict[idx]['공통']['컬럼 한글명']=None # 컬럼 한글명 (정의서 정보 활용 내용 반영 예정)
            self.ResultDict[idx]['공통']['데이터 타입']=data[col].dtypes.name # 데이터 타입
            self.ResultDict[idx]['공통']['null 개수']='{:,}'.format(data[col].isnull().sum()) # null 개수
            self.ResultDict[idx]['공통']['%null']='{:.2%}'.format(data[col].isnull().sum()/data.shape[0]) # %null
            self.ResultDict[idx]['공통']['적재건수']='{:,}'.format(data[col].notnull().sum()) # 적재건수
            self.ResultDict[idx]['공통']['%적재건수']='{:.2%}'.format(data[col].notnull().sum()/data.shape[0]) # %적재건수
        
        # Step 3-2: 연속형 영역 QC 수행
        for idx in self.ResultDict.keys():
            self.ResultDict[idx]['연속형']={}
            col=self.ResultDict[idx]['공통']['컬럼 영문명']
            
            if any(keyword in self.ResultDict[idx]['공통']['데이터 타입'] for keyword in ['float', 'int']):
                
                self.ResultDict[idx]['연속형']['최솟값']=str(data[col].min()) # 최솟값
                self.ResultDict[idx]['연속형']['최댓값']=str(data[col].max()) # 최댓값
                self.ResultDict[idx]['연속형']['평균']=str(data[col].mean()) # 평균
                self.ResultDict[idx]['연속형']['표준편차']=str(data[col].std()) # 표준편차
                self.ResultDict[idx]['연속형']['중위수']=str(np.median(data[col])) # 표준편차
                
            else:
                self.ResultDict[idx]['연속형']['최솟값']=None # 최솟값
                self.ResultDict[idx]['연속형']['최댓값']=None # 최댓값
                self.ResultDict[idx]['연속형']['평균']=None # 평균
                self.ResultDict[idx]['연속형']['표준편차']=None # 표준편차
                self.ResultDict[idx]['연속형']['중위수']=None # 표준편차
        
        # Step 3-3: 범주형 영역 QC 수행
        for idx in self.ResultDict.keys():
            self.ResultDict[idx]['범주형']={}
            col=self.ResultDict[idx]['공통']['컬럼 영문명']
            
            if any(keyword in self.ResultDict[idx]['공통']['데이터 타입'] for keyword in ['object']):
        
                self.ResultDict[idx]['범주형']['범주수']='{:,}'.format(data[col].nunique(dropna=True)) # 범주수
                if data[col].nunique(dropna=True) <= 5:
                    self.ResultDict[idx]['범주형']['범주']=data[col].unique().tolist() # 범주
                    self.ResultDict[idx]['범주형']['%범주']={value_: '{:.2%}'.format((data[col].loc[data[col]==value_].shape[0])/(data.shape[0])) for value_ in data[col].unique().tolist()} # %범주
                else:
                    self.ResultDict[idx]['범주형']['범주']=data[col].unique()[:2].tolist() + ['...'] + data[col].unique()[-2:].tolist() # 범주
                    self.ResultDict[idx]['범주형']['%범주']={value_: '{:.3%}'.format((data[col].loc[data[col]==value_].shape[0])/(data.shape[0])) for value_ in data[col].unique()[:5].tolist()} # %범주
                    self.ResultDict[idx]['범주형']['%범주']['그 외']='{:.3%}'.format((data[col].loc[~(data[col].isin(data[col].unique()[:5].tolist()))].shape[0])/(data.shape[0]))
                self.ResultDict[idx]['범주형']['정의된 범주 외']=None # 정의된 범주 외 (정의서 정보 활용 내용 반영 예정)
                self.ResultDict[idx]['범주형']['정의된 범주 외 수']=None # 정의된 범주 외 수 (정의서 정보 활용 내용 반영 예정)
                if len(data[col].mode(dropna=True).values.tolist()) <= 3:
                    self.ResultDict[idx]['범주형']['최빈값']=data[col].mode(dropna=True).values.tolist() # 최빈값
                    self.ResultDict[idx]['범주형']['최빈값 수']={mode_: '{:,}'.format(data[col].loc[data[col]==mode_].shape[0]) for mode_ in data[col].mode(dropna=True).values.tolist()} # 최빈값 수
                    self.ResultDict[idx]['범주형']['%최빈값']={mode_: '{:.2%}'.format((data[col].loc[data[col]==mode_].shape[0])/(data.shape[0])) for mode_ in data[col].mode(dropna=True).values.tolist()} # %최빈값
                else:
                    self.ResultDict[idx]['범주형']['최빈값']=data[col].mode(dropna=True).values.tolist()[:2] + ['...'] # 최빈값
                    self.ResultDict[idx]['범주형']['최빈값 수']={mode_col: '{:,}'.format(data[col].loc[data[col]==mode_col].shape[0]) for mode_col in data[col].mode(dropna=True).values.tolist()[:2]} # 최빈값 수
                    self.ResultDict[idx]['범주형']['%최빈값']={mode_col: '{:.2%}'.format((data[col].loc[data[col]==mode_col].shape[0])/(data.shape[0])) for mode_col in data[col].mode(dropna=True).values.tolist()[:2]} # %최빈값
                    
            else:
                self.ResultDict[idx]['범주형']['범주수']=None # 범주수
                self.ResultDict[idx]['범주형']['범주']=None # 범주
                self.ResultDict[idx]['범주형']['%범주']=None # %범주
                self.ResultDict[idx]['범주형']['정의된 범주 외']=None # 정의된 범주 외
                self.ResultDict[idx]['범주형']['정의된 범주 외 수']=None # 정의된 범주 외 수
                self.ResultDict[idx]['범주형']['최빈값']=None # 최빈값
                self.ResultDict[idx]['범주형']['최빈값 수']=None # 최빈값 수
                self.ResultDict[idx]['범주형']['%최빈값']=None # %최빈값
        
        
        # Step 3-4: 비고 영역 QC 수행
        for idx in self.ResultDict.keys():
            self.ResultDict[idx]['비고']={}
            col=self.ResultDict[idx]['공통']['컬럼 영문명']
        
            self.ResultDict[idx]['비고']['비고']=None # 컬럼 한글명 (정의서 정보 활용 내용 반영 예정)
            
    def convert_to_richtext(self, src):
        if type(src) is list:
            tgt=',\n'.join(src)[:-1]
        elif type(src) is dict:
            tgt="\n".join("{}: {},".format(k, v) for k, v in src.items())[:-1]
            
        return tgt
        
    def save(self):
        
        # 결과 저장
        
        # Step 1: Json 파일 저장
        with open(os.path.join(self.PATH['OUTPUT'], 'QC결과서.json'), 'w') as f:
            json.dump(self.ResultDict, f, ensure_ascii=False)
        
        # Step 2: QC 결과서 산출물 생성
        # Step 2-1: 기본 Excel 파일 생성
        SubCol1, SubCol2=[], []
        for key1 in self.RelCategory.keys():
            SubCol1+=[key1] * len(self.RelCategory[key1])
            SubCol2+=self.RelCategory[key1]
        
        ColList=[SubCol1, SubCol2]
        OutputPath=os.path.join(self.PATH['OUTPUT'], 'QC결과서.xlsx')
        
        ResultList=[]
        
        for idx in self.ResultDict.keys():
            ResultList_=[]
            for key1 in self.ResultDict[idx]:
                ResultList_+=[self.convert_to_richtext(self.ResultDict[idx][key1][key2]) if ((type(self.ResultDict[idx][key1][key2]) is list) or (type(self.ResultDict[idx][key1][key2]) is dict)) else self.ResultDict[idx][key1][key2] for key2 in self.ResultDict[idx][key1]]
                
            ResultList.append(ResultList_)
        
        ResultDoc=pd.DataFrame(ResultList, columns=ColList)
        ResultDoc.to_excel(OutputPath, sheet_name='Table1', index=True, header=True)
        
        # # 저장한 Excel 파일 편집 
        wb=load_workbook(OutputPath)
        ws=wb.active
        
        ws.delete_rows(3)
        ws.delete_cols(1)

        for mcr in ws.merged_cells:
            if 1 < mcr.min_col:
                mcr.shift(col_shift=-1)
            elif 1 <= mcr.max_col:
                mcr.shrink(right=1)
                
        thin = Side(border_style="thin", color="000000")
        
        for i_, row in enumerate(ws.rows):
            for cell_ in row:
                if i_==0:
                    if cell_.value in ['공통']:
                        cell=ws[cell_.coordinate]
                        cell.fill = PatternFill("solid", fgColor="bfbfbf")
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif cell_.value in ['연속형']:
                        cell=ws[cell_.coordinate]
                        cell.fill = PatternFill("solid", fgColor="f4b084")
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif cell_.value in ['범주형']:
                        cell=ws[cell_.coordinate]
                        cell.fill = PatternFill("solid", fgColor="9bc2e6")
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif cell_.value in ['비고']:
                        ws.merge_cells(f'V1:V2')
                        cell=ws[cell_.coordinate]
                        cell.fill = PatternFill("solid", fgColor="bfbfbf")
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                elif i_==1:
                    if cell_.value in self.RelCategory['공통'] + self.RelCategory['비고']:
                        cell=ws[cell_.coordinate]
                        cell.fill = PatternFill("solid", fgColor="d9d9d9")
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif cell_.value in self.RelCategory['연속형']:
                        cell=ws[cell_.coordinate]
                        cell.fill = PatternFill("solid", fgColor="f8cbad")
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif cell_.value in self.RelCategory['범주형']:
                        cell=ws[cell_.coordinate]
                        cell.fill = PatternFill("solid", fgColor="bdd7ee")
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell=ws[cell_.coordinate]
                    cell.alignment = Alignment(vertical='center', wrap_text=True)
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        
        ColumnDimension(ws, bestFit=True)
        
        wb.save(OutputPath)