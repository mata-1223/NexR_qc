import os, re, json, copy, time, datetime
import numpy as np
import pandas as pd
import pyarrow.parquet as pq
from openpyxl import load_workbook, Workbook

class PreProcess:

    def __init__(self, inputPath, docsPath):
        # 공통 결측치를 설정 
        self.naList = ["?", "na", "null", "Null", "NULL", " ", "[NULL]"]
        addNa = input(f"결측값을 추가할 수 있습니다. \n기본 설정된 결측값: {self.naList} \n추가하고 싶은 결측값을 작성해주십시오:")
        if len(addNa) != 0:
            self.naList += addNa.replace(" ", "").split(sep=",")
            self.naList = list(set(self.naList))
        print(f"현재 공통 결측값 리스트는 {self.naList} 입니다.")
        fileName = inputPath.split("/")[-1]
        self.fileName = re.split(".csv|.parquet", fileName)[0]
        try:
            # 테이블 정의서
            self.tableDocs = pd.read_excel(os.path.join(docsPath, list(file for file in os.listdir(docsPath) if "테이블" in file)[0]), usecols=lambda x: "Unnamed" not in x)
            self.tableDocs.columns = [i.replace("\n","") for i in self.tableDocs.columns]
            # 컬럼 정의서
            self.colDocs = pd.read_excel(os.path.join(docsPath, list(file for file in os.listdir(docsPath) if "컬럼" in file)[0]), usecols=lambda x: "Unnamed" not in x, dtype={"코드대분류\n(그룹코드ID)":"string"})
            self.colDocs.columns = [i.replace("\n","") for i in self.colDocs.columns]
            self.colDocs = pd.merge(self.colDocs, self.tableDocs.loc[:,["시스템명(영문)", "스키마명" , "테이블명(영문)", "DB 유형"]], on=["시스템명(영문)", "스키마명" , "테이블명(영문)"], how="left")
            self.colDocs.loc[:,"데이터타입"] = self.colDocs.loc[:,"데이터타입"].str.upper()
            # 코드 정의서
            self.codeDocs = pd.read_excel(os.path.join(docsPath, list(file for file in os.listdir(docsPath) if "코드" in file)[0]), usecols=lambda x: "Unnamed" not in x, dtype={"코드값":"string", "코드대분류\n(그룹코드ID)":"string"})
            self.codeDocs.columns = [i.replace("\n","") for i in self.codeDocs.columns]
            # 데이터 타입 정의서
            self.dtypeDocs = pd.read_excel(os.path.join(docsPath, list(file for file in os.listdir(docsPath) if "Datatype" in file)[0]), usecols=lambda x: "Unnamed" not in x)
            self.dtypeDocs.loc[:,"DataType"] = self.dtypeDocs.loc[:,"DataType"].str.upper()
            # 데이터 형식 변환
            self.colDocs = pd.merge(self.colDocs, self.dtypeDocs, left_on=["데이터타입","DB 유형"], right_on=["DataType", "DBMS"], how="left")
            self.colDocs = self.colDocs.drop(["NewDataType","DBMS","DataType"], axis=1)
            self.colDocs = self.colDocs.drop_duplicates()
            self.dbType = dict(zip(self.colDocs.loc[self.colDocs.loc[:,"테이블명(영문)"]==self.fileName, "컬럼명"], self.colDocs.loc[self.colDocs.loc[:,"테이블명(영문)"]==self.fileName, "PyDataType"]))

            # data loading
            if ".csv" in inputPath:
                self.data = pd.read_csv(inputPath, dtype={key: value for key, value in self.dbType.items() if value=="string"}, parse_dates=[key for key, value in self.dbType.items() if value=="datetime64"], infer_datetime_format=True, na_values=self.naList)
    #             elif ".parquet" in inputPath:
    #                 self.data = pq.read_pandas(inputPath).to_pandas()
    #                 self.fileName = fileName.split(".parquet")[0]

            self.data = self.data.astype(self.dbType, errors="ignore")
            self.overview = dict()
        except:
            print("해당 파일이 존재하지 않습니다. 경로를 확인하세요.")

    def na_check(self):
        # 공통 결측치를 설정 
        self.naList = ["?", "na", "null", "Null", "NULL", " "]
        addNa = input(
            f"결측값을 추가할 수 있습니다. \n기본 설정된 결측값: {self.naList} \n추가하고 싶은 결측값을 작성해주십시오:"
        )
        if len(addNa) != 0:
            self.naList += addNa.replace(" ", "").split(sep=",")
            self.naList = list(set(self.naList))
        # 결측값을 처리 
        self.data[self.data.isin(self.naList)] = np.nan
    
    def summary(self):
        self.result = dict()
        """
        기초 정보를 담고 있는 Json 파일 생성
        """
        # data의 row, column 수를 확인
        rows, columns = self.data.shape
        # data 내 전체 결측값을 확인
        totalNull = sum(self.data.isnull().sum())
        # 중복되는 row가 있는지 확인
        duplicateRow = sum(self.data.duplicated())
        # 중복 row의 index를 확인
        duplicateIdx = [
            idx
            for idx, result in self.data.duplicated().to_dict().items()
            if result is True
        ]
        # 연속형 변수를 확인
        numericVar = list(self.data.select_dtypes(include=np.number).columns)
        num = dict({"count": len(numericVar), "variables": numericVar})
        # 범주형 변수를 확인
        stringVar = list(self.data.select_dtypes(include="string").columns)
        string = dict({"count": len(stringVar), "variables": stringVar})
        # 범주형 변수를 확인
        timeVar = list(self.data.select_dtypes(include="datetime64").columns)
        time = dict({"count": len(timeVar), "variables": timeVar})
        self.result["overview"] = {
            "dataset":{
                "rows": rows,
                "cols": columns,
                "null": totalNull,
                "null%": round(totalNull / rows, 2),
                "numericVar": num,
                "stringVar": string,
                "duplicateRow": duplicateRow,
                "duplicateRowIdx": duplicateIdx,
            }
        }
        """
        각 변수 summary값 dict 형태로 저장
        """
        self.result["edaResult"] = {"Numeric": dict(), "String": dict()}
        for colName in self.data.columns:
            if colName in self.result["overview"]["dataset"]["numericVar"]["variables"]:
                summary = dict({"korName":self.colDocs.loc[(self.colDocs.loc[:,"테이블명(영문)"]==self.fileName)&(self.colDocs.loc[:,"컬럼명"]==colName),"속성명(컬럼한글명)"].tolist()[0]})
                summaryTmp = self.data.loc[:,colName].describe().fillna(0)
                # json으로 저장하기 위해 형식을 변경한다. 
                for i in summaryTmp.keys():
                    summary[i] = float(summaryTmp[i])
                summary["count"] = len(self.data.loc[:,colName])
                summary["nullCount"] = int(self.data.loc[:,colName].isnull().sum())
                summary["nullProp"] = summary["nullCount"] / len(self.data)
                summary["nullOnly"] = (1 if summary["nullProp"] == 1 else 0)
                self.result["edaResult"]["Numeric"][colName] = dict(summary)
            elif colName in self.result["overview"]["dataset"]["stringVar"]["variables"]:
                summary = dict({"korName":self.colDocs.loc[(self.colDocs.loc[:,"테이블명(영문)"]==self.fileName)&(self.colDocs.loc[:,"컬럼명"]==colName),"속성명(컬럼한글명)"].tolist()[0]})
                summary["PK"] = self.colDocs.loc[(self.colDocs.loc[:,"테이블명(영문)"]==self.fileName)&(self.colDocs.loc[:,"컬럼명"]==colName),"PK여부"].tolist()[0]
                summary["FK"] = self.colDocs.loc[(self.colDocs.loc[:,"테이블명(영문)"]==self.fileName)&(self.colDocs.loc[:,"컬럼명"]==colName),"FK여부"].tolist()[0]
                summary["count"] = len(self.data.loc[:,colName])
                ftable = dict(self.data.loc[:,colName].value_counts())
                ftableProp = dict(self.data.loc[:,colName].value_counts()/len(self.data))
                # json으로 저장하기 위해 형식을 변경한다. 
                for i in ftable.keys():
                    ftable[i] = int(ftable[i])
                for i in ftableProp.keys():
                    ftableProp[i] = float(ftableProp[i])
                codeCheckCat = self.colDocs.loc[(self.colDocs.loc[:,"테이블명(영문)"]==self.fileName)&(self.colDocs.loc[:,"컬럼명"]==colName), "코드대분류(그룹코드ID)"].tolist()[0]
                codeCheckDocs = self.codeDocs[self.codeDocs["코드대분류(그룹코드ID)"]==codeCheckCat]
                summary["classDefined"] = list(codeCheckDocs["코드값"].unique())
                summary["classCount"] = ftable
                summary['classProp'] = ftableProp
                summary["nullCount"] = int(self.data.loc[:,colName].isnull().sum())
                summary["nullProp"] = summary["nullCount"] / len(self.data)
                summary["nullOnly"] = (1 if summary["nullProp"] == 1 else 0)
                self.result["edaResult"]["String"][colName] = dict(summary)
    
    def eda(self):
        """
        total summary table 생성
        """
        # MUlti index 세팅
        totSummaryCol = [
            ["컬럼"] * 3 + ["연속형 대상"] * 7 + ["범주형 대상"] * 4 + ["공통"] * 4,
            ["컬럼명", "한글명", "타입", 
             "최소값", "25%", "50%", "75%", "최대값", "평균", "표준편차",
             "범주 수", "정의된 범주 외", "정의된 범주 외%", "최빈값",
             "NULL값", "NULL수", "NULL%", "적재건수"]
        ]
        # total summary table
        self.result["totalSummary"] = pd.DataFrame(columns=totSummaryCol)
        # each summary table
        self.result["eachSummary"] = dict({"Numeric":dict(), "String":dict()})
        # FK, PK가 아닌 String Var 리스트
        strList = [col for col in self.result["edaResult"]["String"].keys() if (self.result["edaResult"]["String"][col].get("PK")=="N")&(self.result["edaResult"]["String"][col].get("FK")=="N")]
        for colType in self.result["edaResult"].keys():
            for colName in self.result["edaResult"][colType].keys():
                dataSample = self.result["edaResult"][colType][colName]
                if colType == "Numeric":
                    # total summary
                    totSummary = pd.DataFrame(
                        np.array(
                            (
                                colName,
                                dataSample["korName"],
                                colType,
                                round(dataSample["min"], 2),
                                round(dataSample["25%"], 2),
                                round(dataSample["50%"], 2),
                                round(dataSample["75%"], 2),
                                round(dataSample["max"], 2),
                                round(dataSample["mean"], 2),
                                round(dataSample["std"], 2),
                                dataSample["nullCount"],
                                round(dataSample["nullProp"] * 100, 2),
                                dataSample["count"],
                            )
                        ).reshape(1, 13),
                        columns=[
                            ["컬럼"] * 3 + ["연속형 대상"] * 7 + ["공통"] * 3,
                            ["컬럼명", "한글명", "타입", 
                            "최소값", "25%", "50%", "75%", "최대값", "평균", "표준편차", 
                            "NULL수", "NULL%", "적재건수"]
                        ],
                    )
                    # each summary
                    # base
                    eachSummary = pd.DataFrame(self.data.loc[:,colName].describe()).T
                    eachSummary.rename(columns={"count":"빈도수", "mean":"평균", "std":"표준편차", "min":"최소값", "max":"최대값"}, inplace=True)
#                     eachSummary.columns = ["빈도수", "평균", "표준편차", "최소값", "25%", "50%", "75%", "최대값"]
                    eachSummary = eachSummary.loc[:, ["빈도수", "최소값", "25%", "50%", "75%", "최대값", "평균", "표준편차"]]
                    self.result["eachSummary"][colType][colName] = {colName: {"base": eachSummary}}
                    # correlation
                    if (self.result["edaResult"][colType][colName]["nullOnly"] == 0)&(self.data.select_dtypes(include=np.number).shape[1] > 1):
                        self.result["eachSummary"][colType][colName][colName]["correlation"] = self.data.corr()
                    # group by time
                    if (self.result["edaResult"][colType][colName]["nullOnly"] == 0)&(self.data.select_dtypes(include="datetime").shape[1] > 0):
                        data = copy.copy(self.data)
                        for timeCol in self.data.select_dtypes(include="datetime").columns:
                            if self.data.loc[:,timeCol].isnull().sum() != len(self.data):
                                data["Year"] = self.data.loc[:,timeCol].dt.year
                                data["Month"] = self.data.loc[:,timeCol].dt.month
                                data["Day"] = self.data.loc[:,timeCol].dt.day
                                data["Hour"] = self.data.loc[:,timeCol].dt.hour
                                self.result["eachSummary"][colType][colName][timeCol] = dict()
                                for timeFilter in [["Year"], ["Year", "Month"], ["Year", "Month", "Day"], ["Year", "Month", "Day", "Hour"]]:
                                    timeGroupbyData = data[[colName]+[f"{name}" for name in timeFilter]].groupby([f"{name}" for name in timeFilter]).describe()
                                    timeGroupbyData = timeGroupbyData.reset_index(drop=False)
                                    mIndex = [(x, y) for x, y in zip([timeCol]*len(timeFilter)+[colName]*8, timeFilter+["빈도수","평균","표준편차","최소값","25%","50%","75%","최대값"])]
                                    timeGroupbyData.columns = pd.MultiIndex.from_tuples(mIndex)
                                    # column index 정의
                                    timeGroupbyCol = [[timeCol]*len(timeFilter)+[colName]*8, timeFilter+["빈도수", "최소값", "25%", "50%", "75%", "최대값", "평균", "표준편차"]]
                                    self.result["eachSummary"][colType][colName][timeCol]["_".join(timeFilter)] = timeGroupbyData.reindex(columns=timeGroupbyCol)
                    # group by String
                    if (self.result["edaResult"][colType][colName]["nullOnly"] == 0)&(self.data.select_dtypes(include="string").shape[1] > 0):
                        for col in strList:
                            if self.result["edaResult"]["String"][col]["nullOnly"] == 0:
                                strGroupbyData = self.data.loc[:,[colName, col]].groupby(col).describe()
                                mIndex = [(col, "범주")] + strGroupbyData.columns.tolist()
                                strGroupbyData.reset_index(drop=False, inplace=True)
                                strGroupbyData.columns = pd.MultiIndex.from_tuples(mIndex)
                                strGroupbyData.rename(columns={"count":"빈도수", "mean":"평균", "std":"표준편차", "min":"최소값", "max":"최대값"}, level=1, inplace=True)
                                # column index 정의
                                strGroupbyCol = [[col]+[colName]*8, ["범주", "빈도수", "최소값", "25%", "50%", "75%", "최대값", "평균", "표준편차"]]
                                self.result["eachSummary"][colType][colName][col] = dict()
                                self.result["eachSummary"][colType][colName][col]["base"] = strGroupbyData.reindex(columns=strGroupbyCol)
                                # correlation
                                if (self.data.select_dtypes(include=np.number).shape[1] > 1):
                                    self.result["eachSummary"][colType][colName][col]["correlation"] = self.data.groupby(col).corr()
                elif colType == "String":
                    classUndefined = {key: value for key, value in self.result["edaResult"][colType][colName]["classCount"].items() if key not in self.result["edaResult"][colType][colName]["classDefined"]}
                    # total summary
                    totSummary = pd.DataFrame(
                        np.array(
                            (
                                colName,
                                dataSample["korName"],
                                colType,
                                len(dataSample["classCount"]),
                                ", ".join(self.data.loc[:,colName].mode().tolist()) if (len(list(dataSample["classCount"].items()))!=dataSample["count"])&(len(list(dataSample["classCount"].items())) > 0)&(self.result["edaResult"]["String"][colName]["PK"]=="N")&(self.result["edaResult"]["String"][colName]["FK"]=="N") else "",
                                dataSample["nullCount"],
                                round(dataSample["nullCount"] / len(self.data) * 100, 2),
                                dataSample["count"],
                            )
                        ).reshape(1, 8),
                        columns=[
                            ["컬럼"] * 3 + ["범주형 대상"] * 2 + ["공통"] * 3,
                            ["컬럼명", "한글명", "타입",
                             "범주 수", "최빈값",
                             "NULL수", "NULL%", "적재건수"]
                        ],
                    )
                    if (len(self.result["edaResult"][colType][colName]["classDefined"]) > 0)&(len(classUndefined) > 0):
                        totSummary2 = pd.DataFrame(
                            np.array(
                                (
                                    ", ".join(classUndefined.keys()),
                                    round(sum(classUndefined.values()) / len(self.data) * 100, 2)
                                )
                            ).reshape(1,2),
                            columns = [["범주형 대상"] * 2, ["정의된 범주 외", "정의된 범주 외%"]
                            ]
                        )
                        totSummary = pd.concat((totSummary, totSummary2), axis=1)
                    # FK, PK가 아닌 String Var    
                    if colName in strList:
                        # excel file 정의
                        edaResult = copy.copy(self.result["edaResult"]["String"][colName])
                        edaResult.pop("classDefined")
                        freqTable = pd.DataFrame(edaResult) if len(edaResult["classCount"]) > 0 else pd.DataFrame([edaResult])
                        freqTable = freqTable.loc[:, ["korName", "classCount", "count", "classProp", "nullCount", "nullProp", "nullOnly"]]
                        if len(self.result["edaResult"]["String"][colName]["classDefined"])>0 :
                            freqTable.insert(3, "defined", [1 if i in self.result["edaResult"]["String"][colName]["classDefined"] else 0 for i in freqTable.index]) ## 속도 개선 여지 존재
                        if freqTable.nullOnly[0] == 1:
                            freqTable[["classCount", "classProp"]] = np.nan
                        freqTable.index = pd.MultiIndex.from_tuples([i for i in zip(freqTable.korName, freqTable.index)])
                        freqTable.drop("korName", axis=1, inplace=True)
                        freqTable.rename(columns={"defined":"범주 정의 여부", "count":"적재건수", "classCount":"빈도수", "classProp":"비율", "nullCount":"결측치 수", "nullProp":"결측 비율", "nullOnly":"전체 결측 여부"}, inplace=True)
                        self.result["eachSummary"][colType][colName]= {colName: {"base": freqTable}}
                        # group by time
                        if (self.result["edaResult"][colType][colName]["nullOnly"] == 0)&(self.data.select_dtypes(include="datetime").shape[1] > 0):
                            data = copy.copy(self.data)
                            for timeCol in self.data.select_dtypes(include="datetime").columns:
                                if self.data.loc[:,timeCol].isnull().sum() != len(self.data):
                                    data["Year"] = self.data.loc[:,timeCol].dt.year
                                    data["Month"] = self.data.loc[:,timeCol].dt.month
                                    data["Day"] = self.data.loc[:,timeCol].dt.day
                                    data["Hour"] = self.data.loc[:,timeCol].dt.hour
                                    self.result["eachSummary"][colType][colName][timeCol] = dict()
                                    for timeFilter in [["Year"], ["Year", "Month"], ["Year", "Month", "Day"], ["Year", "Month", "Day", "Hour"]]:
                    #                 for timeFilter in [["Year", "Month"]]:
                                        timeGroupbyCount = data[[colName]+[f"{name}" for name in timeFilter]].groupby([f"{name}" for name in timeFilter]+[colName]).size().unstack()
                                        timeGroupbyCount.columns = pd.MultiIndex.from_product([[timeGroupbyCount.columns.name], timeGroupbyCount.columns], names=None)
                                        timeGroupbyMode = data.groupby([f"{name}" for name in timeFilter])[[colName]].agg(pd.Series.mode)
                                        timeGroupbyMode.columns = pd.MultiIndex.from_product([timeGroupbyMode.columns, ["최빈값"]], names=None)
                                        timeGroupbyData = pd.merge(timeGroupbyCount, timeGroupbyMode, left_index=True, right_index=True)
                                        timeGroupbyData.loc[:, (colName, "최빈값")] = [",".join(i) if type(i) != str else i for i in timeGroupbyMode[colName]["최빈값"].tolist()]
                                        mIndex = [(x, y) for x, y in zip([timeCol]*len(timeFilter), timeFilter)] + timeGroupbyData.columns.tolist()
                                        timeGroupbyData = timeGroupbyData.reset_index(drop=False)
                                        timeGroupbyData.columns = pd.MultiIndex.from_tuples(mIndex, names=None)
                                        self.result["eachSummary"][colType][colName][timeCol]["_".join(timeFilter)] = timeGroupbyData
                        # group by String Var
                        strExceptList = copy.copy(strList)
                        strExceptList.remove(colName)
                        for strCol in strExceptList:
                            if (self.result["edaResult"][colType][colName]["nullOnly"] == 0)&(self.result["edaResult"][colType][strCol]["nullOnly"] == 0):
                                strGroupbyCount = pd.crosstab(self.data.loc[:,colName], self.data.loc[:,strCol], margins=True, margins_name="합계")
                                mIndex = [(x, y) for x, y in zip([""] + [strCol]*len(strGroupbyCount.columns.tolist()), [colName] + strGroupbyCount.columns.tolist())]
                                strGroupbyCount = strGroupbyCount.reset_index(drop=False)
                                strGroupbyCount.columns = pd.MultiIndex.from_tuples(mIndex)
                                self.result["eachSummary"][colType][colName][strCol] = dict()
                                self.result["eachSummary"][colType][colName][strCol]["base"] = strGroupbyCount
                self.result["totalSummary"] = pd.concat([self.result["totalSummary"], totSummary], ignore_index=True).reindex(columns=totSummaryCol)
    
    def save(self, outputPath, **kwargs):
        saveTime = datetime.datetime.fromtimestamp(time.time()).strftime("%Y%m%d%H%M")
        dbName = self.tableDocs.loc[self.tableDocs["테이블명(영문)"]==self.fileName, "스키마명"].tolist()[0]
        savePath = f"{outputPath}/{dbName}_{self.fileName}_{saveTime}"
        # initial Excel 파일 생성
        if not os.path.exists(savePath):
            # 폴더 생성
            os.makedirs(savePath)
            os.makedirs(os.path.join(savePath, "Numeric"))
            os.makedirs(os.path.join(savePath, "String"))
            # Json 파일 저장
            json.dump(self.result["overview"], open(f"{os.path.join(savePath)}/overview.json", "w"))
            json.dump(self.result["edaResult"], open(f"{os.path.join(savePath)}/edaResult.json", "w"))
            # total summary 저장
            self.result["totalSummary"].to_excel(f"{os.path.join(savePath)}/total_summary.xlsx", engine="xlsxwriter")
            lineDel = dict()
            # each summary 저장
            for colType in self.result["eachSummary"].keys():
                lineDel[colType] = dict()
                for colName in self.result["eachSummary"][colType].keys():
                    lineDel[colType][colName] = dict()
                    # save excel file
                    xlsxWriter = pd.ExcelWriter(f"{os.path.join(savePath, colType, colName)}.xlsx", engine="xlsxwriter")
                    for sheetName in self.result["eachSummary"][colType][colName].keys():
                        lineDel[colType][colName][sheetName] = dict()
                        writeRow = 0
                        delRow = [writeRow+3]
                        if sheetName not in self.data.select_dtypes(include="datetime64").columns:
                            for key in self.result["eachSummary"][colType][colName][sheetName].keys():
                                self.result["eachSummary"][colType][colName][sheetName][key].to_excel(xlsxWriter, sheet_name=sheetName, encoding="utf-8-sig", startrow = writeRow)
                                if colName == sheetName:
                                    writeRow += len(self.result["eachSummary"][colType][colName][sheetName][key]) + 2
                                    delCol = {}
                                else: 
                                    writeRow += len(self.result["eachSummary"][colType][colName][sheetName][key]) + 4
                                    delCol = [1]
                                delRow += [writeRow+3]
                        else:
                            for timeFilter in self.result["eachSummary"][colType][colName][sheetName].keys():
                                self.result["eachSummary"][colType][colName][sheetName][timeFilter].to_excel(xlsxWriter, sheet_name=sheetName, encoding="utf-8-sig", startrow = writeRow)
                                writeRow += len(self.result["eachSummary"][colType][colName][sheetName][timeFilter]) + 4
                                delCol = [1]
                                delRow += [writeRow+3]
                        lineDel[colType][colName][sheetName]["col"] = delCol
                        lineDel[colType][colName][sheetName]["row"] = delRow
                    xlsxWriter.save()

            # 다시 엑셀 포멧팅 작업
            for fType in ["Numeric", "String"]:
                fDir = f"{savePath}/{fType}"
                fList = os.listdir(fDir)
                fList = [x for x in fList if ".xlsx" in x]
                for fName in fList:
                    workBook = load_workbook(f"{fDir}/{fName}")
                    for sheetName in workBook.sheetnames:
                        if sheetName != fName[:(len(fName)-5)]:
                            colDel = lineDel[fType][fName[:(len(fName)-5)]][sheetName]["col"]
                            rowDel = lineDel[fType][fName[:(len(fName)-5)]][sheetName]["row"]
                            if len(colDel) > 0:
                                for col in colDel:
                                    self.modify_cell(col=col, sheet=workBook[sheetName])
                            if len(rowDel) > 0:
                                rowEdit = 0
                                for row in rowDel:
                                    self.modify_cell(row=row+rowEdit, sheet=workBook[sheetName])
                                    rowEdit += -1
                    workBook.save(f"{fDir}/{fName}")
                    
        else:
            print("지정된 저장폴더가 이미 존재합니다.")
            raise SystemExit
            
    def modify_cell(self, **kwargs):
        if kwargs.get("col") is not None:
            kwargs["sheet"].delete_cols(kwargs["col"])
            for mcr in kwargs["sheet"].merged_cells:
                if kwargs.get("col") < mcr.min_col:
                    mcr.shift(col_shift=-1)
                elif kwargs.get("col") <= mcr.max_col:
                    mcr.shrink(right=1)
        if kwargs.get("row") is not None:
            kwargs["sheet"].delete_rows(kwargs["row"])
            for mcr in kwargs["sheet"].merged_cells:
                if kwargs.get("row") < mcr.min_row:
                    mcr.shift(row_shift=-1)
                elif kwargs.get("row") <= mcr.max_row:
                    mcr.shrink(bottom=1)