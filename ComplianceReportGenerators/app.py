from .conf import config
from .filehandle import fhandle
from .dbConnection import dbConnection

import datetime
from dateutil.relativedelta import relativedelta
import os
import pandas as pd
import numpy as np
pd.options.mode.chained_assignment = None

class compliance:
    
    def __init__(self, RawDataInfo, StorageInfo, PreReport, VIPList, EscRaw):
        filehandler = fhandle()
        self.db = dbConnection("mssql")
        self.outFile = filehandler.createOutputDir()
        self.fileDate = filehandler.fileDate
        self.now = filehandler.now
        self.rawDataInfo = pd.read_excel(RawDataInfo)
        storageInfo = pd.ExcelFile(StorageInfo)
        self.storageInfo = storageInfo.parse('R115_Storage_Utilization___Topo',skiprows=7, usecols="A:U")
        self.preReport = pd.ExcelFile(PreReport)
        self.escRaw = pd.read_excel(EscRaw)
        self.vipList = pd.read_excel(VIPList)
        # Run Topology and get High here
        self.Topology()
        #self.highLevelReport = pd.ExcelFile(ASCR) 
        #self.high = self.highLevelReport.parse("Highlevelreport")
        #self.run = Run
    
    #*************************************************************************************************************************************************************
    def Topology(self):
        # Initiating and parsing Excel
        df = self.rawDataInfo
        Run=df["Run"][0]
        self.run = Run
        # Deleting unwanted columns
        del df["Run"]
        del df["PrevEC"]
        del df["EC"]
        # Replace Nan with 0
        df = df.fillna(0)
        try:
            db_engine= self.db
            connection = db_engine.con.connect()
            ## For Topology 3 
            no = "'"+str(self.run)+"'"
            query = "Select * from [dbo].[R158 Shell-Site Collection list] where run="+no
            #query = "select * FROM r158_shellsite where run="+no
            qry = connection.execute(query)
            table = qry.fetchall()
            head = qry.keys()
            n=lambda x: x.upper() if isinstance(x, str) else x
            head=list(map(n, head))
        except Exception as e:
            return "Topology error: "+str(e)
        DFrame = pd.DataFrame(table, columns=head)
        df.insert(7,"Topology_Level_3_Business_Unit",None,allow_duplicates=False)
        df.insert(8,"Topology_Level_4_Business_Unit",None,allow_duplicates=False)
        for i in range(len(df)):
            topo3 = DFrame[DFrame["SITE_COLLECTION_URL"] == df.iloc[i]["SiteName"]]
            if topo3.empty == False:
                df.at[i,"Topology_Level_3_Business_Unit"]=list(topo3["TOPOLOGY_LEVEL_3_BUSINESS_UNIT"])[0]
            else:
                df.at[i,"Topology_Level_3_Business_Unit"]=None
        if connection.closed == False:
            connection.close()
        ## For Topology 4
        df_Storage = self.storageInfo
        for i in range(len(df)):
            topo4 = df_Storage.loc[df_Storage["Site Collection URL"] == df.iloc[i]["SiteName"]]        
            if topo4.empty == False:
                level = topo4["Topology Hierarchy"].to_list()[0]
                if level != "Unknown":
                    if level!=None:
                        if "\\" in level:
                            level = level.split('\\')
                            if len(level) > 4:
                                df.at[i,"Topology_Level_4_Business_Unit"]=level[3]
                            else:
                                df.at[i,"Topology_Level_4_Business_Unit"]="Not Available"
            else:
                df.at[i,"Topology_Level_4_Business_Unit"]=None
        ## Overview Part
        # Read from previous reportin
        # previous report excel input by user
        # read overview sheet
        # get Present data from df
        current = self.now
        ad = relativedelta(months=-1)
        ColDate = current + ad
        ColDate = ColDate.strftime("%b-%y")
        PrevReportData = self.preReport
        overview = PrevReportData.parse('Overview')
        colLen = len(overview.columns)
        confidential = int(df["Confidential"].sum())
        MC = int(df["MC"].sum())
        EU_C = int((df[df['SiteName'].str.contains('EU00')])['Confidential'].sum())
        EU_MC= int((df[df['SiteName'].str.contains('EU00')])['MC'].sum())
        US_C = int((df[df['SiteName'].str.contains('US00')])['Confidential'].sum())
        US_MC = int((df[df['SiteName'].str.contains('US00')])['MC'].sum())
        NGA_C = int((df[df['SiteName'].str.contains('NGA00')])['Confidential'].sum())
        NGA_MC = int((df[df['SiteName'].str.contains('NGA00')])['MC'].sum())
        My_C = int((df[df['SiteName'].str.contains('MY00')])['Confidential'].sum())
        My_MC = int((df[df['SiteName'].str.contains('MY00')])['MC'].sum())
        Chn_C = int((df[df['SiteName'].str.contains('CHN00')])['Confidential'].sum())
        Chn_MC = int((df[df['SiteName'].str.contains('CHN00')])['MC'].sum())
        head = []
        LastMonth = overview
        overview.insert(colLen-1, ColDate,"")
        if(overview.iloc[0,0] == "# Confidential"):
            overview.at[0, ColDate]=confidential
        if(overview.iloc[1,0] == "# Most Confidential"):
            overview.at[1, ColDate]=MC
        for i in range(len(overview)):
            if(overview.iloc[i,0] == "Central Environment"):
                if(overview.iloc[i+1,0] == "# Confidential"):
                    overview.at[i+1,ColDate]=EU_C
                    overview.at[i+2,ColDate]=EU_MC
            if(overview.iloc[i,0] == "China"):
                if(overview.iloc[i+1,0] == "# Confidential"):
                    overview.at[i+1,ColDate]=Chn_C
                    overview.at[i+2,ColDate]=Chn_MC
            if(overview.iloc[i,0] == "CBJ"):
                if(overview.iloc[i+1,0] == "# Confidential"):
                    overview.at[i+1,ColDate]=My_C
                    overview.at[i+2,ColDate]=My_MC
            if(overview.iloc[i,0] == "Nigeria"):
                if(overview.iloc[i+1,0] == "# Confidential"):
                    overview.at[i+1,ColDate]=NGA_C
                    overview.at[i+2,ColDate]=NGA_MC
        for colName in list(overview.columns):
            if isinstance(colName, datetime.date):
                head.append(colName.strftime("%b-%y"))
            else:
                head.append(colName)
        overview_Data=overview.values.tolist()
        overview_format=pd.DataFrame(overview_Data, columns=head)
        for i in range(len(overview_format)):
            last = overview_format.columns[-2]
            PrevLast = overview_format.columns[-3]
            if overview_format[last][i] != "nan" and overview_format[last][i] !=None and overview_format[last][i] != '':
                cur = overview_format[last][i]
                pre = overview_format[PrevLast][i]
                diff = cur - pre
                overview_format.at[i, 'Delta']=int(diff)
        #output 
        outFile = self.outFile+'\\ASCR '+self.fileDate+'.xlsx'
        outFileName = 'ASCR '+self.fileDate+'.xlsx'
        outWriter = pd.ExcelWriter(outFile, engine = 'xlsxwriter')
        with pd.ExcelWriter(outFile) as writer:
            df.to_excel(writer, sheet_name="Highlevelreport",index = False)
            overview_format.to_excel(writer, sheet_name="Overview",index = False)
        # Check if output is present
        self.high = df
        return outFileName

    
    #*************************************************************************************************************************************************************
    def Escalation(self):
        # Initiating and parsing Excel
        allSites = self.escRaw
        # Deleting unwanted columns
        del allSites["RecurringEC"]
        # Replace Nan with 0
        allSites.update(allSites[['LineManager']].fillna('NULL'))
        allSites.update(allSites[['RecurringConfidential','RecurringMC']].fillna(0))
        ## Recurring MC
        MC = allSites[(allSites["RecurringMC"]!=0)]
        del MC["RecurringConfidential"]
        ## Recurring Confidential
        RC = allSites[(allSites["RecurringConfidential"]!=0)]
        del RC["RecurringMC"]
        #output 
        outFile = self.outFile+'\\EscalationReport '+self.fileDate+'.xlsx'
        outFileName = 'EscalationReport '+self.fileDate+'.xlsx'
        with pd.ExcelWriter(outFile) as writer:
            allSites.to_excel(writer, sheet_name="ALLSITES",index = False)
            MC.to_excel(writer, sheet_name="MC",index = False)
            RC.to_excel(writer, sheet_name="Confidential",index = False)
        return outFileName
    
    def DS(self):
        df = self.high
        high = df[df['Topology_Level_2_Business']=='Downstream']
        high = high[["SiteName","Confidential","MC","Site_Collection_Name","Topology_Level_2_Business","Topology_Level_3_Business_Unit","Topology_Level_4_Business_Unit","Primary_Site_Owner","Secondary_Site_Owner","AllConfidential","AllDocument"]]
        high["VIP"]="NO"
        high.reset_index(inplace=True,drop=True)
        # VIPList
        vip = self.vipList
        for i in range(len(high)):
            ifvip = vip["VIP"+vip['Email Address']==high.iloc[i]["Primary_Site_Owner"]]
            if ifvip.empty == False:
                high["Primary_Site_Owner"][i] = ifvip["Email Address"].values[0]
                high["VIP"][i] = "Yes"
        # Downstream IT
        ds = df[df['Topology_Level_4_Business_Unit']=='Downstream IT']
        ds = ds[["SiteName","Confidential","MC","Site_Collection_Name","Topology_Level_2_Business","Topology_Level_3_Business_Unit","Topology_Level_4_Business_Unit","Primary_Site_Owner","Secondary_Site_Owner","AllConfidential","AllDocument"]]
        ds["VIP"]="No"	
        ds.reset_index(inplace=True, drop=True)
        #ds["VIP"] = np.where(ds['Primary_Site_Owner'].str.contains("VIP"), True, False)
        for i in range(len(ds)):
            IfVip = vip["VIP"+vip['Email Address']==ds.iloc[i]["Primary_Site_Owner"]]
            if IfVip.empty == False:
                ds.at[i, "Primary_Site_Owner"] = ifvip["Email Address"].values[0]
                ds.at[i, "VIP"]="YES"
        # Output Section 
        outFile = self.outFile+'\\ASCR '+self.fileDate+' DS.xlsx'
        outFileName = 'ASCR '+self.fileDate+' DS.xlsx'
        with pd.ExcelWriter(outFile) as writer:
            high.to_excel(writer, sheet_name="DS",index = False)
            ds.to_excel(writer, sheet_name="DSIT",index = False)
        return outFileName
    
        
    def UP(self):
        # Highlevelreport
        high = self.high
        high = high[high['Topology_Level_2_Business']=='Upstream']
        high = high[["SiteName","Confidential","MC","Site_Collection_Name","Topology_Level_2_Business","Topology_Level_3_Business_Unit","Topology_Level_4_Business_Unit","Primary_Site_Owner","Secondary_Site_Owner","AllConfidential","AllDocument"]]
        high["VIP"]="NO"
        high.reset_index(inplace=True, drop=True)
		# VIPList
        vip = self.vipList
        for i in range(len(high)):
            ifvip = vip["VIP"+vip['Email Address']==high.iloc[i]["Primary_Site_Owner"]]
            if ifvip.empty == False:
                high["Primary_Site_Owner"][i] = ifvip["Email Address"].values[0]
                high["VIP"][i] = "Yes"
        # Output Section 
        outFile = self.outFile+'\\ASCR '+self.fileDate+' UP.xlsx'
        outFileName = 'ASCR '+self.fileDate+' UP.xlsx'
        with pd.ExcelWriter(outFile) as writer:
            high.to_excel(writer, sheet_name="UP",index = False)
        # Check if output is present
        return outFileName
        
    def PT(self):
        high = self.high
        high = high[high['Topology_Level_2_Business']=='Projects and Technology']
        high = high[["SiteName","Confidential","MC","Site_Collection_Name","Topology_Level_2_Business","Topology_Level_3_Business_Unit","Primary_Site_Owner","Secondary_Site_Owner","AllConfidential","AllDocument"]]
        high.reset_index(inplace=True,drop=True)
        # Output Section 
        outFile = self.outFile+'\\ASCR '+self.fileDate+' PT.xlsx'
        outFileName = 'ASCR '+self.fileDate+' PT.xlsx'
        with pd.ExcelWriter(outFile) as writer:
            high.to_excel(writer, sheet_name="PT",index = False)
        # Check if output is present
        return outFileName
        
    def IGNE(self):
        high = self.high
        high = high[high['Topology_Level_2_Business']=='Integrated Gas & New Energies']
        high = high[["SiteName","Confidential","MC","Site_Collection_Name","Topology_Level_2_Business","Topology_Level_3_Business_Unit","Topology_Level_4_Business_Unit","Primary_Site_Owner","Secondary_Site_Owner","AllConfidential","AllDocument"]]
        high["VIP"]="NO"
        high.reset_index(inplace=True,drop=True)
        # VIPList
        vip = self.vipList
        for i in range(len(high)):
            ifvip = vip["VIP"+vip['Email Address']==high.iloc[i]["Primary_Site_Owner"]]
            if ifvip.empty == False:
                high["Primary_Site_Owner"][i] = ifvip["Email Address"].values[0]
                high["VIP"][i] = "Yes"
        outFile = self.outFile+'\\ASCR '+self.fileDate+' IG&NE.xlsx'
        outFileName = 'ASCR '+self.fileDate+' IG&NE.xlsx'
        with pd.ExcelWriter(outFile) as writer:
            high.to_excel(writer, sheet_name="IG&NE",index = False)
        # Check if output is present
        return outFileName
        
    def Global_Finance(self):
        try:
            db_engine = self.db 
            connection = db_engine.con.connect()
            ## Total Site Collections 
            no = "'"+str(self.run)+"'"
            query = "select * FROM [dbo].[R158 Shell-Site Collection list] where run="+no+" and Topology_Level_2_Business='Global Functions' and Topology_Level_3_Business_Unit='Finance'"
            qry = connection.execute(query)
            table = qry.fetchall()
            head = qry.keys()
        except Exception as e:
            return "Global_Finance error: "+str(e)
        Total_SiteCollection = pd.DataFrame(table, columns=head)
        ## Non Compliant Site Collection
        high = self.high
        NonComplaint_siteCollection = high[(high['Topology_Level_2_Business']=='Global Functions') & (high['Topology_Level_3_Business_Unit']=='Finance')]
        del NonComplaint_siteCollection["PrevConfidential"]
        del NonComplaint_siteCollection["PrevMC"]
        del NonComplaint_siteCollection["Secondary_Site_Owner"]
        del NonComplaint_siteCollection["AllConfidential"]
        del NonComplaint_siteCollection["AllDocument"]
        ## Non complaint docurl
        SiteName = NonComplaint_siteCollection["SiteName"]
        SiteName = ",".join('\'' + item + '\'' for item in SiteName.to_list())
        nonUrl_Query = "select SiteUrl, Owner, DocumentUrl, DocId, Category from [dbo].[ASCR_IncompliantDocUrls] where run="+no+"and SiteUrl in ("+SiteName+")"
        #nonUrl_Query = "select SiteUrl, Owner, DocumentUrl, DocId, Category from [dbo].[ASCR_IncompliantDocUrls] where run="+no+"and SiteUrl in (%s)" % ','.join('?' * len(SiteName))
        try:
            nonUrl_qry = connection.execute(nonUrl_Query)
            nonUrl_table = nonUrl_qry.fetchall()
            nonUrl_head = nonUrl_qry.keys()
        except Exception as e:
            return "Global_Finance error: "+str(e)
        NonComplaint_Docurl = pd.DataFrame(nonUrl_table, columns = nonUrl_head)
        if connection.closed == False:
            connection.close()
        # Output Section 
        outFile = self.outFile+'\\ASCR '+self.fileDate+' Global Finance.xlsx'
        outFileName = 'ASCR '+self.fileDate+' Global Finance.xlsx'
        with pd.ExcelWriter(outFile) as writer:
            Total_SiteCollection.to_excel(writer, sheet_name="Total SiteCollections",index = False)
            NonComplaint_siteCollection.to_excel(writer, sheet_name="NonComplaint siteCollection",index = False)
            NonComplaint_Docurl.to_excel(writer, sheet_name="NonComplaint Docurl",index = False)
        # Check if output is present
        return outFileName
        
    def Global_Function(self):
        # Highlevelreport
        high = self.high
        high = high[high['Topology_Level_2_Business']=='Global Functions']
        high = high[["SiteName","Confidential","MC","Site_Collection_Name","Topology_Level_2_Business","Topology_Level_3_Business_Unit","Topology_Level_4_Business_Unit","Primary_Site_Owner","Secondary_Site_Owner","AllConfidential","AllDocument"]]
        high["VIP"]="NO"
        high.reset_index(inplace=True, drop=True)
        # VIPList
        vip = self.vipList
        for i in range(len(high)):
            ifvip = vip["VIP"+vip['Email Address']==high.iloc[i]["Primary_Site_Owner"]]
            if ifvip.empty == False:
                high["Primary_Site_Owner"][i] = ifvip["Email Address"].values[0]
                high["VIP"][i] = "Yes"
        #Output Section
        outFile = self.outFile+'\\ASCR '+self.fileDate+' GF.xlsx'
        outFileName = 'ASCR '+self.fileDate+' GF.xlsx'
        with pd.ExcelWriter(outFile) as writer:
            high.to_excel(writer, sheet_name="GF",index = False)
        # Check if output is present
        return outFileName

if __name__ == "__main__":
    a = compliance()
    b = a.G
       

    