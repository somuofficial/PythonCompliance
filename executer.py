from ComplianceReportGenerators import compliance, FileCheck
import sys
#RawDataInfo, StorageInfo, PreReport, VIPList):
if (len(sys.argv) == 6):
    RawDataInfo = sys.argv[1]
    StorageInfo = sys.argv[2]
    PreReport = sys.argv[3]
    VIPList = sys.argv[4]
    EscRaw = sys.argv[5]
    print("Initiating Excel Formatter")
    rawDataInfo = FileCheck(RawDataInfo)
    storageInfo = FileCheck(StorageInfo)
    preReport = FileCheck(PreReport)
    vipList = FileCheck(VIPList)
    escRaw = FileCheck(EscRaw)
    if rawDataInfo.out:
        if storageInfo.out:
            if preReport.out:
                if vipList.out:
                    if escRaw.out:
                        exe = compliance(rawDataInfo.filePath, storageInfo.filePath, preReport.filePath, vipList.filePath, escRaw.filePath)
                        ds = exe.DS()
                        print(ds)
                        up = exe.UP()
                        print(up)
                        pt = exe.PT()
                        print(pt)
                        igne = exe.IGNE()
                        print(igne)
                        gf = exe.Global_Function()
                        print(gf)
                        esc = exe.Escalation()
                        print(esc)
                        gfin = exe.Global_Finance()
                        print(gfin)
                        print("Reports generated in: "+ exe.outFile)
                    else:
                        print('Escalation raw file not present'+ escRaw.filePath)
                else:
                    print('VIP list excel not present: '+ vipList.filePath)
            else:
                print('Previous month ASCR report not present: '+ preReport.filePath)
        else:
            print('Storage file not present: '+ storageInfo.filePath)
    else:
        print('Raw data file not present: '+ rawDataInfo.filePath)
else:
    RawDataInfo = input("Insert RawDataInfo data excel file name: ")
    StorageInfo = input("Insert StorageInfo excel file name: ")
    PreReport = input("Insert PreReport excel file name: ")
    VIPList = input("Insert VIP List excel file name: ")
    EscRaw = input("Insert Escalation raw file name: ")
    print("Initiating Excel Formatter")
    rawDataInfo = FileCheck(RawDataInfo)
    storageInfo = FileCheck(StorageInfo)
    preReport = FileCheck(PreReport)
    vipList = FileCheck(VIPList)
    escRaw = FileCheck(EscRaw)
    if rawDataInfo.out:
        if storageInfo.out:
            if preReport.out:
                if vipList.out:
                    if escRaw.out:
                        exe = compliance(rawDataInfo.filePath, storageInfo.filePath, preReport.filePath, vipList.filePath, escRaw.filePath)
                        ds = exe.DS()
                        print(ds)
                        up = exe.UP()
                        print(up)
                        pt = exe.PT()
                        print(pt)
                        igne = exe.IGNE()
                        print(igne)
                        gf = exe.Global_Function()
                        print(gf)
                        esc = exe.Escalation()
                        print(esc)
                        gfin = exe.Global_Finance()
                        print(gfin)
                        print("Reports generated in: "+ exe.outFile)
                    else:
                        print('Escalation raw file not present'+ escRaw.filePath)
                else:
                    print('VIP list excel not present: '+ vipList.filePath)
            else:
                print('Previous month ASCR report not present: '+ preReport.filePath)
        else:
            print('Storage file not present: '+ storageInfo.filePath)
    else:
        print('Raw data file not present: '+ rawDataInfo.filePath)