from QualityCheck import *

PathDict = {}
PathDict["ROOT"] = os.getcwd()
PathDict["DATA"] = os.path.join(PathDict["ROOT"], "data")

DataDict = {}
for path in os.listdir(PathDict["DATA"]):
    if not path.startswith("."):
        data_name = os.path.splitext(os.path.basename(path))[0].upper()
        DataDict[data_name] = pd.read_csv(os.path.join(PathDict["DATA"], path))

if __name__ == "__main__":

    Process = QualityCheck(DataDict)
    Process.data_check()
    Process.document_check()
    Process.na_check()
    Process.run()
    Process.save()
