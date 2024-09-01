from TermSheetGenerator import *

termSheetDict = {
    "SDAC":SDACGenerator
}

def createInstance(termSheetType,termPath=None):
    if termSheetType in termSheetDict:
        return termSheetDict[termSheetType]() if not termPath else termSheetDict[termSheetType](termPath)
    else:
        raise Exception(f"Termsheet Type: {termSheetType} is not supported")

if __name__ == "__main__":
    termSheet = "SDAC"
    tsGenerator = createInstance(termSheet)
    tsGenerator.generateNewTermSheet("test")