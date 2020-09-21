# Author: Tim Gowan, FWRI
# Date: Feb 2011
# Purpose: QA/QC raw data tables from FWRI team.  Will update existing fields, create new fields, and flag invalid values
# Requirements: raw data tables from FWRI team which exist in SQL database

##################   Declare  Variables #######################################
output = "C:\\Documents and Settings\\tim.gowan\\Desktop\\SEWSoutput.txt" #path and name of output file to be created which will store results
SQLDB = "QAQC_1011"  #Name of SQL database
DDSOURCE = "FLA"    #Name for DDSOURCE, according to survey team
prefix = "f1"       #prefix for FILEID, according to survey type and team
year1 = 2010       #years included in the season
year2 = 2011
minalt = 60         #limits for altitude of plane (in meters)
maxalt = 750
minlat = 25         #limits for LATDEG and LONGDEG (note: LONGDEG negative in West)
maxlat = 35
minlong = -82
maxlong = -75
maxtime = 0.25      #maximum time allowed between events, in decimal hours (e.g. 0.25 = 15 minutes)
minspeed = 12.5     #limits for plane speed (in meters/s)
maxspeed = 103
maxdescent = -4.57  #limits for change in altitude of plane (in meters/s)
maxascent = 4.57
############################################################################################################

# Import system modules
import sys, string, os, math
import win32com.client
import time
import datetime
from pyproj import Geod # Module for geodetic distance calculations
D = Geod(ellps='GRS80') # Use GRS80 ellipsoid (NAD 83) for geodetic calculations

CPU_time1 = time.time()  # Start time (value to estimate total processing time)

#Create and open the output text file to store results - 'w' for writing (an existing file with the same name will be erased), 'a' opens the file for appending, 'r' for read only
f = open(output, 'w')

#Print today's date to the output file for your record
now = datetime.datetime.now()
print >> f, now.strftime("%m-%d-%Y") #print >> f, will print the statement to the output txt file 

# Connect to the SQL database 
SQL_DSN = 'PROVIDER=sqloledb;DATA SOURCE=.\SQLEXPRESS;Initial Catalog=' + SQLDB + ';Integrated Security=SSPI;'    # Connection string
SQLconn = win32com.client.Dispatch(r'ADODB.Connection')
SQLconn.Open(SQL_DSN) # Connect to the SQL server
############################### for test runs only
try:
    SQLconn.Execute("DROP TABLE SEWS_tables")
    print 'Deleted existing SEWS_tables'
except:
    print 'SEWS_tables has yet to be created' #delete list of SEWS tables if it already exists
#############################
SQLconn.Execute("SELECT name INTO SEWS_tables FROM sys.Tables WHERE name LIKE '%SEWS%' ORDER BY name") #Create new table with list of all SEWS tables
SQLrs = win32com.client.Dispatch(r'ADODB.Recordset')
SQLrs.Open('SELECT * FROM SEWS_tables ORDER BY name', SQLconn, 1, 3)   # Open list of tables as a recordset
Ntables = SQLrs.RecordCount
print >> f, 'There are ' + str(Ntables - 1) + ' SEWS tables.' #Ntables - 1 = all tables except "SEWS_tables"

i = 0
SQLrs.MoveFirst() #Move to first record (first table in list)

# Loop through each SEWS table, passing over "SEWS_tables"
#####change to < Ntables in actual run to check all tables
while i < Ntables: #while there are still tables to check
    stringname = SQLrs.Fields.Item('name').Value
    suffix = "tables"
    if stringname.endswith(suffix): #pass over "SEWS_tables"
        pass
    else:
        # Begin checking data table
        Table = SQLrs.Fields.Item('name').Value     #Table is all data from a single survey day
        ############################### for test runs only
        try:
            SQLconn.Execute("DROP TABLE " + Table + "Check")
            print "Deleted existing " + Table + "Check"
        except:
            print "CheckTable " + Table + "has yet to be created" #delete test table if it already exists
        SQLconn.Execute("SELECT * INTO " + Table + "Check FROM " + Table + " ORDER BY EVENTNO") #Copy data into checking table
        #############################
        print >> f, "\nBegin checking " + Table #print statement to the output txt file (\n puts statement on next line in text file)
        # Add, update and alter fields
        SQLconn.Execute("sp_RENAME '" + Table + "Check.DATE_', 'DATE', 'COLUMN'") #Rename fields with extra underscore
        SQLconn.Execute("sp_RENAME '" + Table + "Check.TIME_', 'TIME', 'COLUMN'")
        SQLconn.Execute("sp_RENAME '" + Table + "Check.LONG_', 'LONG', 'COLUMN'")
        SQLconn.Execute("sp_RENAME '" + Table + "Check.NUMBER_', 'NUMBER', 'COLUMN'")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ADD LATDEG int, LATMIN float, LONGDEG int, LONGMIN float") #add deg and min fields, declaring data types
        SQLconn.Execute("UPDATE " + Table + "Check SET LATDEG=FLOOR(LAT)") #calculate deg and min fields
        SQLconn.Execute("UPDATE " + Table + "Check SET LONGDEG=CEILING(LONG)")
        SQLconn.Execute("UPDATE " + Table + "Check SET LATMIN=(LAT-LATDEG)*60")
        SQLconn.Execute("UPDATE " + Table + "Check SET LONGMIN=(LONG-LONGDEG)*-60")
        SQLconn.Execute("UPDATE " + Table + "Check SET SIGHTNO=0 WHERE SIGHTNO IS NULL") #Change SIGHTNO, NUMBER, and NUMCALF null values to 0
        SQLconn.Execute("UPDATE " + Table + "Check SET SIGHTNO=0 WHERE SPECCODE IS NULL")
        SQLconn.Execute("UPDATE " + Table + "Check SET NUMBER=0 WHERE NUMBER IS NULL")
        SQLconn.Execute("UPDATE " + Table + "Check SET NUMCALF=0 WHERE NUMCALF IS NULL")
        SQLconn.Execute("UPDATE " + Table + "Check SET LEGSTAGE=NULL WHERE LEGSTAGE=0")
        SQLconn.Execute("UPDATE " + Table + "Check SET PHOTOS=1 WHERE PHOTOS IS NULL AND SIGHTNO <> 0") #Set PHOTOS to 1 (no photos) when there is a sighting but PHOTOS is null
        SQLconn.Execute("sp_RENAME '" + Table + "Check.TIME', 'GPSTIME', 'COLUMN'") #Rename TIME field
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN SPECCODE varchar(4)") #Change data types
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN GPSTIME int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN WX varchar(1)")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN EVENTNO int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN SIGHTNO int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN ANHEAD int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN BEAUFORT int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN CLOUD int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN CONFIDNC int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN IDREL int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN PHOTOS int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN VISIBLTY float")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN GLAREL int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN GLARER int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ALTER COLUMN NUMBER int")
        SQLconn.Execute("ALTER TABLE " + Table + "Check ADD FILEID varchar(8), TIME int, DDSOURCE varchar(4), LENGTH_m float, TIMELENGTH_dh float, SPEED_ms float, ALTCHANGE_ms float") #add fields, declaring data types
        SQLconn.Execute("ALTER TABLE " + Table + "Check ADD MONTH int, DAY int, YEAR int") #FWRI does not include Month, Day, or Year - add fields, declaring data types
        SQLconn.Execute("UPDATE " + Table + "Check SET DDSOURCE = '" + DDSOURCE + "'")
        SQLconn.Execute("UPDATE " + Table + "Check SET TIME = GPSTIME - 50000") #convert GPS time to local time

        # Add remaining BEHAV fields if necessary; set data type to integer
        k = 1
        while k < 16: #should be BEHAV1-BEHAV15
            Behav = str("BEHAV" + str(k))
            SQLconn.Execute("if exists (select COLUMN_NAME from INFORMATION_SCHEMA.columns where table_name = '" + Table + "Check' and column_name = '" + Behav + "') alter table " + Table + "Check ALTER COLUMN " + Behav + " int") #if the field already exists, set data type as integer
            SQLconn.Execute("if not exists (select COLUMN_NAME from INFORMATION_SCHEMA.columns where table_name = '" + Table + "Check' and column_name = '" + Behav + "') alter table " + Table + "Check add " + Behav + " int") #if it doesn't exist, add new field
            k = k + 1
             
        Tablers = win32com.client.Dispatch(r'ADODB.Recordset')
        Tablers.Open("SELECT * FROM " + Table + "Check ORDER BY EVENTNO, SIGHTNO", SQLconn, 1, 3)   # Open table as a recordset
        Nrecords = Tablers.RecordCount #count number of records in table
        

        PrevFileID = 'Start'
        PrevEVENTNO = 'Start'
        PrevTime = 'Start'
        PrevSIGHTNO = 'Start'
        PrevBeaufort = 'Start'
        PrevCloud = 'Start'
        PrevLegtype = 'Start'
        PrevLegstage = 'Start'
        PrevGlareL = 'Start'
        PrevGlareR = 'Start'
        PrevVisiblty = 'Start'
        PrevWX = 'Start'
        PrevHour = 'Start'
        PrevLong = 'Start'
        PrevLat = 'Start'
        PrevAlt = 'Start'
        j = 0
        Tablers.MoveFirst() #move to first record in table
        while j < Nrecords: ######Change to Nrecords to search all records
            StringDate = str(Tablers.Fields.Item('DATE').Value) #convert value of DATE as a string
            Date = time.strptime(StringDate, "%m/%d/%Y ") #parse DATE into components
            year = str(Date[0]) #year is first component of parsed Date
            idyear = year[2:4]
            julian = str(Date[7]) #julian date is 8th component of parsed Date

            #Set MONTH, DAY and YEAR
            Tablers.Fields.Item('MONTH').Value = Date[1]
            Tablers.Fields.Item('DAY').Value = Date[2]
            Tablers.Fields.Item('YEAR').Value = Date[0]
            
            #Define Hour, Min, Seconds from TIME field
            TimeString = str(Tablers.Fields.Item('GPSTIME').Value)
            if len(TimeString) == 6:   # for double digit hours
                PMTimehour = TimeString[0:2]
                if int(PMTimehour) > 12:
                    Timehour = str(int(PMTimehour) - 12)
                elif int(PMTimehour) < 13:
                    Timehour = PMTimehour
                Timeminute = TimeString[2:4]
                Timeseconds = int(TimeString[4:6])
            elif len(TimeString) == 5:    #for single digit hours
                PMTimehour = TimeString[0:1]
                Timehour = TimeString[0:1]
                Timeminute = TimeString[1:3]
                Timeseconds = int(TimeString[3:5])
            else:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid Time"

            #populate FILEID field
            if Date[7] < 10:
                fileid = prefix + idyear + "00" + julian  #pad julian dates with extra 0's if < 100
            elif (Date[7] > 9) and (Date[7] < 100):
                fileid = prefix + idyear + "0" + julian
            elif Date[7] > 99:
                fileid = prefix + idyear + julian
            Tablers.Fields.Item('FILEID').Value = fileid
            if PrevFileID == 'Start':                   #when on the first record in the table, print the survey date and FILEID to the output file
                print >> f, "-Survey date: " + StringDate + ", fileID: " + str(fileid)

            #Check FileID matches previous record; checks date as well since FileID populated from Date field
            if PrevFileID == 'Start':  #pass if first record in table
                pass
            elif Tablers.Fields.Item('FILEID').Value <> PrevFileID:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid FILEID: " + str(Tablers.Fields.Item('FILEID').Value)

            #Check components of Date are valid and match Time
            if Date[0] <> year1 and Date[0] <> year2:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid year"
            if Date[1] < 1 or Date[1] > 12:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid month"
            if Date[2] < 1 or Date[2] > 31:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid day"
            CurrentTimeTime = int(Timehour) + float(Timeminute)/60 + float(Timeseconds)/3600   #calculate times as decimal hours
            if Timeseconds < 0 or Timeseconds > 59:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid seconds"
            if PrevTime == 'Start':
                pass
            elif Tablers.Fields.Item('TIME').Value < PrevTime:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Time decreasing"

            #if EVENTNO is duplicate it should have different SIGHTNO but same environmental data (eg. TIME) as previous record  
            if Tablers.Fields.Item('EVENTNO').Value == PrevEVENTNO and Tablers.Fields.Item('SIGHTNO').Value <= PrevSIGHTNO:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid EVENTNO"
            if Tablers.Fields.Item('EVENTNO').Value == PrevEVENTNO and Tablers.Fields.Item('TIME').Value <> PrevTime:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Duplicate EVENTNO"

            #Check SIGHTNO increases
            if PrevSIGHTNO == 'Start' or Tablers.Fields.Item('SIGHTNO').Value == 0:  #pass if first record in table or no sighting
                pass
            elif Tablers.Fields.Item('SIGHTNO').Value <= PrevSIGHTNO:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Duplicate SIGHTNO"

            #Check ALT within range
            if Tablers.Fields.Item('ALT').Value is None:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Missing ALT"
            elif Tablers.Fields.Item('ALT').Value is not None and (float(Tablers.Fields.Item('ALT').Value) < minalt or float(Tablers.Fields.Item('ALT').Value) > maxalt):
                AltFt = round(Tablers.Fields.Item('ALT').Value*3.2808399, 1) #Convert ALT to feet, rounded to 1 decimal place for output
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid ALT: " + str(AltFt) + " ft"

            #Check ANHEAD - not valid without sighting; range=0-17,21,22
            if (Tablers.Fields.Item('SIGHTNO').Value == 0) and (Tablers.Fields.Item('ANHEAD').Value is not None):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - ANHEAD without SIGHTNO"
            if (Tablers.Fields.Item('ANHEAD').Value is not None) and (Tablers.Fields.Item('ANHEAD').Value < 0 or (Tablers.Fields.Item('ANHEAD').Value > 17 and Tablers.Fields.Item('ANHEAD').Value <> 21 and Tablers.Fields.Item('ANHEAD').Value <> 22)):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid ANHEAD"

            #Set and check Beaufort - set as last value when null; range=0-6,7,9
            if (Tablers.Fields.Item('BEAUFORT').Value is None) and (PrevBeaufort == 'Start'):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - First BEAUFORT is null. Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            elif Tablers.Fields.Item('BEAUFORT').Value is None and (PrevBeaufort <> 'Start'):
                Tablers.Fields.Item('BEAUFORT').Value = PrevBeaufort
            if Tablers.Fields.Item('BEAUFORT').Value < 0 or (Tablers.Fields.Item('BEAUFORT').Value > 6 and Tablers.Fields.Item('BEAUFORT').Value <> 7 and Tablers.Fields.Item('BEAUFORT').Value <> 9):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid BEAUFORT"
            
            #Loop to check all Behav fields - not valid without sighting; range=00-30,34-38,40-48,50-55,58-72,75-94,97-98
            m = 1
            while m < 16:
                Behav = str("BEHAV" + str(m))
                if (Tablers.Fields.Item(Behav).Value is not None) and (Tablers.Fields.Item('SIGHTNO').Value == 0):
                    print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - BEHAV without SIGHTNO"
                elif (Tablers.Fields.Item(Behav).Value is not None) and ((Tablers.Fields.Item(Behav).Value < 0) or (Tablers.Fields.Item(Behav).Value > 98) or (Tablers.Fields.Item(Behav).Value > 30 and Tablers.Fields.Item(Behav).Value < 34) \
                                                                         or (Tablers.Fields.Item(Behav).Value > 38 and Tablers.Fields.Item(Behav).Value < 40) or (Tablers.Fields.Item(Behav).Value > 48 and Tablers.Fields.Item(Behav).Value < 50) \
                                                                         or (Tablers.Fields.Item(Behav).Value > 55 and Tablers.Fields.Item(Behav).Value < 58) or (Tablers.Fields.Item(Behav).Value > 72 and Tablers.Fields.Item(Behav).Value < 75) or (Tablers.Fields.Item(Behav).Value > 94 and Tablers.Fields.Item(Behav).Value < 97)):
                    print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid " + Behav

                m = m + 1

            #Check Cloud - range=0,1-4,9
            if (Tablers.Fields.Item('CLOUD').Value is None) and (PrevCloud == 'Start'):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - First CLOUD is null. Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            elif Tablers.Fields.Item('CLOUD').Value is None and (PrevCloud <> 'Start'):
                Tablers.Fields.Item('CLOUD').Value = PrevCloud
            if Tablers.Fields.Item('CLOUD').Value < 0 or (Tablers.Fields.Item('CLOUD').Value > 4 and Tablers.Fields.Item('CLOUD').Value <> 9):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid CLOUD"

            #Check CONFIDNC - not valid without sighting, required for sightings, range=00-11   
            if (Tablers.Fields.Item('SIGHTNO').Value == 0) and (Tablers.Fields.Item('CONFIDNC').Value is not None):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - CONFIDNC without SIGHTNO"
            if (Tablers.Fields.Item('SIGHTNO').Value <> 0) and (Tablers.Fields.Item('CONFIDNC').Value is None):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - SIGHTNO without CONFIDNC"
            if (Tablers.Fields.Item('CONFIDNC').Value is not None) and (Tablers.Fields.Item('CONFIDNC').Value < 0 or Tablers.Fields.Item('CONFIDNC').Value > 11):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid CONFIDNC"
            # where NUMBER is high and CONFIDNC is low (i.e. strong), change CONFIDNC to 10 (no estimate of confidence level)
            if Tablers.Fields.Item('CONFIDNC').Value is not None and (Tablers.Fields.Item('NUMBER').Value > 5 and Tablers.Fields.Item('CONFIDNC').Value == 0):
                Tablers.Fields.Item('CONFIDNC').Value = 10
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Changed CONFIDNC to 10"
            if Tablers.Fields.Item('CONFIDNC').Value is not None and (Tablers.Fields.Item('NUMBER').Value > 10 and Tablers.Fields.Item('CONFIDNC').Value < 3):
                Tablers.Fields.Item('CONFIDNC').Value = 10
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Changed CONFIDNC to 10"

            #Check and set LEGTYPE
            if (Tablers.Fields.Item('LEGTYPE').Value is None) and (PrevLegtype == 'Start'):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - First LEGTYPE is null. Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            elif Tablers.Fields.Item('LEGTYPE').Value is None and (PrevLegtype <> 'Start'):
                Tablers.Fields.Item('LEGTYPE').Value = PrevLegtype
            if Tablers.Fields.Item('LEGTYPE').Value <> 9:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid LEGTYPE"
            
            #Set LEGSTAGE where null
            if (Tablers.Fields.Item('LEGSTAGE').Value is None) and (PrevLegstage == 'Start'):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - First LEGSTAGE is null. Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            elif (Tablers.Fields.Item('LEGSTAGE').Value is None) and (PrevLegstage <> 5 and PrevLegstage <> 1 and PrevLegstage <> 'Start'):
                Tablers.Fields.Item('LEGSTAGE').Value = PrevLegstage
            elif (Tablers.Fields.Item('LEGSTAGE').Value is None) and (PrevLegstage == 1):
                Tablers.Fields.Item('LEGSTAGE').Value = 2
            elif (Tablers.Fields.Item('LEGSTAGE').Value is None) and (PrevLegstage == 5): #LEGSTAGE only remains Null between LEGSTAGEs 5 and 1
                Tablers.Fields.Item('LEGSTAGE').Value is None
            #Check LEGSTAGE
            #LEGSTAGE for first event should = 1
            if PrevLegstage == 'Start' and (Tablers.Fields.Item('LEGSTAGE').Value is None or Tablers.Fields.Item('LEGSTAGE').Value <> 1):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - First LEGSTAGE not 1. Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            elif (Tablers.Fields.Item('LEGSTAGE').Value is not None and Tablers.Fields.Item('LEGSTAGE').Value == 1) and (PrevLegstage <> 'Start' and PrevLegstage <> 5):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - LEGSTAGE out of order"
            #range = 0/null,1,2,5; should not decrease except for LEGSTAGE 1
            if Tablers.Fields.Item('LEGSTAGE').Value is not None and (Tablers.Fields.Item('LEGSTAGE').Value <> 0 and Tablers.Fields.Item('LEGSTAGE').Value <> 1 and Tablers.Fields.Item('LEGSTAGE').Value <> 2 and Tablers.Fields.Item('LEGSTAGE').Value <> 5):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid LEGSTAGE"
            if (Tablers.Fields.Item('LEGSTAGE').Value is not None and Tablers.Fields.Item('LEGSTAGE').Value <> 1) and Tablers.Fields.Item('LEGSTAGE').Value < PrevLegstage:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - LEGSTAGE out of order"
            
            #Set GLAREL and GLARER - should be done after LEGSTAGE is set
            if (Tablers.Fields.Item('GLAREL').Value is None) and (Tablers.Fields.Item('LEGSTAGE').Value is not None):
                Tablers.Fields.Item('GLAREL').Value = PrevGlareL
            if (Tablers.Fields.Item('GLARER').Value is None) and (Tablers.Fields.Item('LEGSTAGE').Value is not None):
                Tablers.Fields.Item('GLARER').Value = PrevGlareR
            #Check GLAREL and GLARER - range=0-3
            if Tablers.Fields.Item('GLAREL').Value is not None and (Tablers.Fields.Item('GLAREL').Value < 0 or Tablers.Fields.Item('GLAREL').Value > 3):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid GLAREL"
            if Tablers.Fields.Item('GLARER').Value is not None and (Tablers.Fields.Item('GLARER').Value < 0 or Tablers.Fields.Item('GLARER').Value > 3):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid GLARER"

            #Check plane HEADING - range=0-360
            if Tablers.Fields.Item('HEADING').Value is not None and (float(Tablers.Fields.Item('HEADING').Value) < 0 or float(Tablers.Fields.Item('HEADING').Value) > 360):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid HEADING"

            #Check IDREL - not valid without sighting, required for sightings, range=1-3,9
            if (Tablers.Fields.Item('SIGHTNO').Value == 0) and (Tablers.Fields.Item('IDREL').Value is not None):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - IDREL without SIGHTNO"
            if (Tablers.Fields.Item('SIGHTNO').Value <> 0) and (Tablers.Fields.Item('IDREL').Value is None):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - SIGHTNO without IDREL"
            if Tablers.Fields.Item('IDREL').Value is not None and (Tablers.Fields.Item('IDREL').Value < 1 or (Tablers.Fields.Item('IDREL').Value > 3 and Tablers.Fields.Item('IDREL').Value <> 9)):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid IDREL"

            #Check LATDEG and LONGDEG
            if Tablers.Fields.Item('LATDEG').Value is None or (Tablers.Fields.Item('LATDEG').Value < minlat or Tablers.Fields.Item('LATDEG').Value > maxlat):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid LATDEG"
            if Tablers.Fields.Item('LONGDEG').Value is None or (Tablers.Fields.Item('LONGDEG').Value < minlong or Tablers.Fields.Item('LONGDEG').Value > maxlong):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid LONGDEG"

            #Check LATMIN and LONGMIN - range=0-60
            if Tablers.Fields.Item('LATMIN').Value is None or (Tablers.Fields.Item('LATMIN').Value < 0 or Tablers.Fields.Item('LATMIN').Value > 60):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid LATMIN"
            if Tablers.Fields.Item('LONGMIN').Value is None or (Tablers.Fields.Item('LONGMIN').Value < 0 or Tablers.Fields.Item('LONGMIN').Value > 60):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid LONGMIN"

            #Check NUMBER - not valid without sighting, required for sightings unless CONFIDNC=11
            if (Tablers.Fields.Item('SIGHTNO').Value == 0) and (Tablers.Fields.Item('NUMBER').Value <> 0):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - NUMBER without SIGHTNO"
            if (Tablers.Fields.Item('SIGHTNO').Value <> 0) and (Tablers.Fields.Item('NUMBER').Value is None or Tablers.Fields.Item('NUMBER').Value == 0) and (Tablers.Fields.Item('CONFIDNC').Value <> 11):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - SIGHTNO without NUMBER"
            if Tablers.Fields.Item('NUMBER').Value > 50:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - High NUMBER"

            #Check NUMCALF - not valid without sighting, must be < NUMBER
            if (Tablers.Fields.Item('SIGHTNO').Value == 0) and (Tablers.Fields.Item('NUMCALF').Value <> 0):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - NUMCALF without SIGHTNO"
            if (Tablers.Fields.Item('NUMCALF').Value > 0) and (Tablers.Fields.Item('NUMCALF').Value >= Tablers.Fields.Item('NUMBER').Value):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - NUMCALF >= NUMBER"

            #Check PHOTOS - not valid without sighting, required for sightings, range=1-5
            if (Tablers.Fields.Item('SIGHTNO').Value == 0) and (Tablers.Fields.Item('PHOTOS').Value is not None):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - PHOTOS without SIGHTNO"
            if (Tablers.Fields.Item('SIGHTNO').Value <> 0) and (Tablers.Fields.Item('PHOTOS').Value is None or Tablers.Fields.Item('PHOTOS').Value == 0):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - SIGHTNO without PHOTOS"
            if Tablers.Fields.Item('PHOTOS').Value is not None and (Tablers.Fields.Item('PHOTOS').Value < 1 or Tablers.Fields.Item('PHOTOS').Value > 5):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid PHOTOS"

            #Check SPECCODE - not valid without sighting, required for sightings, must be 4 characters and an allowable code
            if (Tablers.Fields.Item('SIGHTNO').Value == 0) and (Tablers.Fields.Item('SPECCODE').Value is not None):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - SPECCODE without SIGHTNO"
            if (Tablers.Fields.Item('SIGHTNO').Value <> 0) and (Tablers.Fields.Item('SPECCODE').Value is None):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - SIGHTNO without SPECCODE"
            if Tablers.Fields.Item('SPECCODE').Value is not None and (len(str(Tablers.Fields.Item('SPECCODE').Value)) <> 4):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid SPECCODE"
            if Tablers.Fields.Item('SPECCODE').Value is not None and (Tablers.Fields.Item('SPECCODE').Value <> "AMAL" and  Tablers.Fields.Item('SPECCODE').Value <> "ANSH" and  Tablers.Fields.Item('SPECCODE').Value <> "ASDO" and  Tablers.Fields.Item('SPECCODE').Value <> "BASH" and  Tablers.Fields.Item('SPECCODE').Value <> "BELU" and  Tablers.Fields.Item('SPECCODE').Value <> "BEWH" and  Tablers.Fields.Item('SPECCODE').Value <> "BFTU" and  Tablers.Fields.Item('SPECCODE').Value <> "BLBW" and  Tablers.Fields.Item('SPECCODE').Value <> "BLSH" and  Tablers.Fields.Item('SPECCODE').Value <> "BLWH" and  Tablers.Fields.Item('SPECCODE').Value <> "BODO" and  Tablers.Fields.Item('SPECCODE').Value <> "BRWH" and  Tablers.Fields.Item('SPECCODE').Value <> "CLDO" and  Tablers.Fields.Item('SPECCODE').Value <> "CNRA" and  Tablers.Fields.Item('SPECCODE').Value <> "DSWH" and  Tablers.Fields.Item('SPECCODE').Value <> "FIWH" and  Tablers.Fields.Item('SPECCODE').Value <> "FKWH" and  Tablers.Fields.Item('SPECCODE').Value <> "FLFI" and  Tablers.Fields.Item('SPECCODE').Value <> "FRDO" and  Tablers.Fields.Item('SPECCODE').Value <> "GEBW" and  Tablers.Fields.Item('SPECCODE').Value <> "GOBW" and  Tablers.Fields.Item('SPECCODE').Value <> "GRAM" and  Tablers.Fields.Item('SPECCODE').Value <> "GRSE" and  Tablers.Fields.Item('SPECCODE').Value <> "GRTU" and  Tablers.Fields.Item('SPECCODE').Value <> "GRWH" and  Tablers.Fields.Item('SPECCODE').Value <> "HAPO" and  Tablers.Fields.Item('SPECCODE').Value <> "HASE" and  Tablers.Fields.Item('SPECCODE').Value <> "HATU" and  Tablers.Fields.Item('SPECCODE').Value <> "HGSE" and  Tablers.Fields.Item('SPECCODE').Value <> "HHSH" and  Tablers.Fields.Item('SPECCODE').Value <> "HOSE" and  Tablers.Fields.Item('SPECCODE').Value <> "HPSE" and  Tablers.Fields.Item('SPECCODE').Value <> "HUWH" and  Tablers.Fields.Item('SPECCODE').Value <> "KIWH" and  Tablers.Fields.Item('SPECCODE').Value <> "LETU" and  Tablers.Fields.Item('SPECCODE').Value <> "LFPW" and  Tablers.Fields.Item('SPECCODE').Value <> "LOTU" and  Tablers.Fields.Item('SPECCODE').Value <> "MANA" and  Tablers.Fields.Item('SPECCODE').Value <> "MARA" and  Tablers.Fields.Item('SPECCODE').Value <> "MHWH" and  Tablers.Fields.Item('SPECCODE').Value <> "MIWH" and  Tablers.Fields.Item('SPECCODE').Value <> "NBWH" and  Tablers.Fields.Item('SPECCODE').Value <> "OBDO" and  Tablers.Fields.Item('SPECCODE').Value <> "OCSU" and  Tablers.Fields.Item('SPECCODE').Value <> "ORTU" and  Tablers.Fields.Item('SPECCODE').Value <> "OTBI" and  Tablers.Fields.Item('SPECCODE').Value <> "PIWH" and  Tablers.Fields.Item('SPECCODE').Value <> "POBE" and  Tablers.Fields.Item('SPECCODE').Value <> "PSDO" and  Tablers.Fields.Item('SPECCODE').Value <> "PSWH" and  Tablers.Fields.Item('SPECCODE').Value <> "PYKW" and  Tablers.Fields.Item('SPECCODE').Value <> "RITU" and  Tablers.Fields.Item('SPECCODE').Value <> "RIWH" and  Tablers.Fields.Item('SPECCODE').Value <> "RTDO" and  Tablers.Fields.Item('SPECCODE').Value <> "SADO" and  Tablers.Fields.Item('SPECCODE').Value <> "SCFI" and  Tablers.Fields.Item('SPECCODE').Value <> "SCRA" and  Tablers.Fields.Item('SPECCODE').Value <> "SEWH" and  Tablers.Fields.Item('SPECCODE').Value <> "SFPW" and  Tablers.Fields.Item('SPECCODE').Value <> "SNDO" and  Tablers.Fields.Item('SPECCODE').Value <> "SOBW" and  Tablers.Fields.Item('SPECCODE').Value <> "SPDO" and  Tablers.Fields.Item('SPECCODE').Value <> "SPWH" and  Tablers.Fields.Item('SPECCODE').Value <> "STDO" and  Tablers.Fields.Item('SPECCODE').Value <> "SWFI" and  Tablers.Fields.Item('SPECCODE').Value <> "TRBW" and  Tablers.Fields.Item('SPECCODE').Value <> "TUNS" and  Tablers.Fields.Item('SPECCODE').Value <> "UNBA" and  Tablers.Fields.Item('SPECCODE').Value <> "UNBF" and  Tablers.Fields.Item('SPECCODE').Value <> "UNBS" and  Tablers.Fields.Item('SPECCODE').Value <> "UNBW" and  Tablers.Fields.Item('SPECCODE').Value <> "UNCW" and  Tablers.Fields.Item('SPECCODE').Value <> "UNDO" and  Tablers.Fields.Item('SPECCODE').Value <> "UNFI" and  Tablers.Fields.Item('SPECCODE').Value <> "UNFS" and  Tablers.Fields.Item('SPECCODE').Value <> "UNGD" and  Tablers.Fields.Item('SPECCODE').Value <> "UNID" and  Tablers.Fields.Item('SPECCODE').Value <> "UNKO" and  Tablers.Fields.Item('SPECCODE').Value <> "UNLD" and  Tablers.Fields.Item('SPECCODE').Value <> "UNLW" and  Tablers.Fields.Item('SPECCODE').Value <> "UNMW" and  Tablers.Fields.Item('SPECCODE').Value <> "UNRA" and  Tablers.Fields.Item('SPECCODE').Value <> "UNRO" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSB" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSE" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSH" and  Tablers.Fields.Item('SPECCODE').Value <> "UNST" and  Tablers.Fields.Item('SPECCODE').Value <> "UNTU" and  Tablers.Fields.Item('SPECCODE').Value <> "UNWH" and  Tablers.Fields.Item('SPECCODE').Value <> "WBDO" and  Tablers.Fields.Item('SPECCODE').Value <> "WHSH" and  Tablers.Fields.Item('SPECCODE').Value <> "WSDO" and  Tablers.Fields.Item('SPECCODE').Value <> "WTSH" and  Tablers.Fields.Item('SPECCODE').Value <> "ZOOP" and  Tablers.Fields.Item('SPECCODE').Value <> "CG-B" and  Tablers.Fields.Item('SPECCODE').Value <> "CG-C" and  Tablers.Fields.Item('SPECCODE').Value <> "CG-U" and  Tablers.Fields.Item('SPECCODE').Value <> "CRSH" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-B" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-F" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-G" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-J" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-O" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-P" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-R" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-S" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-U" and  Tablers.Fields.Item('SPECCODE').Value <> "DE-W" and  Tablers.Fields.Item('SPECCODE').Value <> "DR-D" and  Tablers.Fields.Item('SPECCODE').Value <> "DR-T" and  Tablers.Fields.Item('SPECCODE').Value <> "DR-U" and  Tablers.Fields.Item('SPECCODE').Value <> "DR-W" and  Tablers.Fields.Item('SPECCODE').Value <> "EXPL" and  Tablers.Fields.Item('SPECCODE').Value <> "FE-H" and  Tablers.Fields.Item('SPECCODE').Value <> "FE-S" and  Tablers.Fields.Item('SPECCODE').Value <> "FE-U" and  Tablers.Fields.Item('SPECCODE').Value <> "FG-A" and  Tablers.Fields.Item('SPECCODE').Value <> "FG-C" and  Tablers.Fields.Item('SPECCODE').Value <> "FG-D" and  Tablers.Fields.Item('SPECCODE').Value <> "FG-G" and  Tablers.Fields.Item('SPECCODE').Value <> "FG-I" and  Tablers.Fields.Item('SPECCODE').Value <> "FG-L" and  Tablers.Fields.Item('SPECCODE').Value <> "FG-O" and  Tablers.Fields.Item('SPECCODE').Value <> "FG-U" and  Tablers.Fields.Item('SPECCODE').Value <> "FRNT" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-C" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-D" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-F" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-G" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-H" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-L" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-P" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-S" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-T" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-U" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-W" and  Tablers.Fields.Item('SPECCODE').Value <> "FV-Z" and  Tablers.Fields.Item('SPECCODE').Value <> "HELO" and  Tablers.Fields.Item('SPECCODE').Value <> "JETS" and  Tablers.Fields.Item('SPECCODE').Value <> "KAYK" and  Tablers.Fields.Item('SPECCODE').Value <> "MV-B" and  Tablers.Fields.Item('SPECCODE').Value <> "MV-C" and  Tablers.Fields.Item('SPECCODE').Value <> "MV-L" and  Tablers.Fields.Item('SPECCODE').Value <> "MV-O" and  Tablers.Fields.Item('SPECCODE').Value <> "MV-S" and  Tablers.Fields.Item('SPECCODE').Value <> "MV-T" and  Tablers.Fields.Item('SPECCODE').Value <> "MV-U" and  Tablers.Fields.Item('SPECCODE').Value <> "MY-L" and  Tablers.Fields.Item('SPECCODE').Value <> "MY-S" and  Tablers.Fields.Item('SPECCODE').Value <> "NV-L" and  Tablers.Fields.Item('SPECCODE').Value <> "NV-S" and  Tablers.Fields.Item('SPECCODE').Value <> "NV-U" and  Tablers.Fields.Item('SPECCODE').Value <> "OI-D" and  Tablers.Fields.Item('SPECCODE').Value <> "OI-L" and  Tablers.Fields.Item('SPECCODE').Value <> "OI-P" and  Tablers.Fields.Item('SPECCODE').Value <> "OI-S" and  Tablers.Fields.Item('SPECCODE').Value <> "PIBO" and  Tablers.Fields.Item('SPECCODE').Value <> "RECV" and  Tablers.Fields.Item('SPECCODE').Value <> "RV-L" and  Tablers.Fields.Item('SPECCODE').Value <> "RV-S" and  Tablers.Fields.Item('SPECCODE').Value <> "RV-U" and  Tablers.Fields.Item('SPECCODE').Value <> "SPFV" and  Tablers.Fields.Item('SPECCODE').Value <> "SV-L" and  Tablers.Fields.Item('SPECCODE').Value <> "SV-S" and  Tablers.Fields.Item('SPECCODE').Value <> "SV-U" and  Tablers.Fields.Item('SPECCODE').Value <> "UNVE" and  Tablers.Fields.Item('SPECCODE').Value <> "WHAL" and  Tablers.Fields.Item('SPECCODE').Value <> "ABDU" and  Tablers.Fields.Item('SPECCODE').Value <> "ARTE" and  Tablers.Fields.Item('SPECCODE').Value <> "ATBR" and  Tablers.Fields.Item('SPECCODE').Value <> "ATPU" and  Tablers.Fields.Item('SPECCODE').Value <> "AUSH" and  Tablers.Fields.Item('SPECCODE').Value <> "BBPL" and  Tablers.Fields.Item('SPECCODE').Value <> "BCPE" and  Tablers.Fields.Item('SPECCODE').Value <> "BHGU" and  Tablers.Fields.Item('SPECCODE').Value <> "BLGU" and  Tablers.Fields.Item('SPECCODE').Value <> "BLKI" and  Tablers.Fields.Item('SPECCODE').Value <> "BLSC" and  Tablers.Fields.Item('SPECCODE').Value <> "BLTE" and  Tablers.Fields.Item('SPECCODE').Value <> "BOGU" and  Tablers.Fields.Item('SPECCODE').Value <> "BRNO" and  Tablers.Fields.Item('SPECCODE').Value <> "BRPE" and  Tablers.Fields.Item('SPECCODE').Value <> "BRTE" and  Tablers.Fields.Item('SPECCODE').Value <> "BSTP" and  Tablers.Fields.Item('SPECCODE').Value <> "BUFF" and  Tablers.Fields.Item('SPECCODE').Value <> "CAGO" and  Tablers.Fields.Item('SPECCODE').Value <> "CATE" and  Tablers.Fields.Item('SPECCODE').Value <> "COEI" and  Tablers.Fields.Item('SPECCODE').Value <> "COLO" and  Tablers.Fields.Item('SPECCODE').Value <> "COMU" and  Tablers.Fields.Item('SPECCODE').Value <> "COSH" and  Tablers.Fields.Item('SPECCODE').Value <> "COTE" and  Tablers.Fields.Item('SPECCODE').Value <> "DCCO" and  Tablers.Fields.Item('SPECCODE').Value <> "DOVE" and  Tablers.Fields.Item('SPECCODE').Value <> "DOWI" and  Tablers.Fields.Item('SPECCODE').Value <> "FOTE" and  Tablers.Fields.Item('SPECCODE').Value <> "GBBG" and  Tablers.Fields.Item('SPECCODE').Value <> "GLGU" and  Tablers.Fields.Item('SPECCODE').Value <> "GRCO" and  Tablers.Fields.Item('SPECCODE').Value <> "GRSC" and Tablers.Fields.Item('SPECCODE').Value <> "GRSH" and  Tablers.Fields.Item('SPECCODE').Value <> "GRSK" and  Tablers.Fields.Item('SPECCODE').Value <> "HERG" and  Tablers.Fields.Item('SPECCODE').Value <> "HOGR" and  Tablers.Fields.Item('SPECCODE').Value <> "ICGU" and  Tablers.Fields.Item('SPECCODE').Value <> "LAGU" and  Tablers.Fields.Item('SPECCODE').Value <> "LBBG" and  Tablers.Fields.Item('SPECCODE').Value <> "LESP" and  Tablers.Fields.Item('SPECCODE').Value <> "LETE" and  Tablers.Fields.Item('SPECCODE').Value <> "LIGU" and  Tablers.Fields.Item('SPECCODE').Value <> "LTDU" and  Tablers.Fields.Item('SPECCODE').Value <> "LTJA" and  Tablers.Fields.Item('SPECCODE').Value <> "MABO" and  Tablers.Fields.Item('SPECCODE').Value <> "MAGW" and  Tablers.Fields.Item('SPECCODE').Value <> "MALL" and  Tablers.Fields.Item('SPECCODE').Value <> "MASH" and  Tablers.Fields.Item('SPECCODE').Value <> "NOFU" and  Tablers.Fields.Item('SPECCODE').Value <> "NOGA" and  Tablers.Fields.Item('SPECCODE').Value <> "PAJA" and  Tablers.Fields.Item('SPECCODE').Value <> "PEEP" and  Tablers.Fields.Item('SPECCODE').Value <> "POJA" and  Tablers.Fields.Item('SPECCODE').Value <> "RAZO" and  Tablers.Fields.Item('SPECCODE').Value <> "RBGU" and  Tablers.Fields.Item('SPECCODE').Value <> "RBME" and  Tablers.Fields.Item('SPECCODE').Value <> "RBTR" and  Tablers.Fields.Item('SPECCODE').Value <> "REKN" and  Tablers.Fields.Item('SPECCODE').Value <> "REPH" and  Tablers.Fields.Item('SPECCODE').Value <> "RNPH" and  Tablers.Fields.Item('SPECCODE').Value <> "ROST" and  Tablers.Fields.Item('SPECCODE').Value <> "ROYT" and  Tablers.Fields.Item('SPECCODE').Value <> "RTLO" and  Tablers.Fields.Item('SPECCODE').Value <> "RUTS" and  Tablers.Fields.Item('SPECCODE').Value <> "SAGU" and  Tablers.Fields.Item('SPECCODE').Value <> "SATE" and  Tablers.Fields.Item('SPECCODE').Value <> "SOSH" and  Tablers.Fields.Item('SPECCODE').Value <> "SOTE" and  Tablers.Fields.Item('SPECCODE').Value <> "SPPL" and  Tablers.Fields.Item('SPECCODE').Value <> "SPSK" and  Tablers.Fields.Item('SPECCODE').Value <> "SUSC" and  Tablers.Fields.Item('SPECCODE').Value <> "TBMU" and  Tablers.Fields.Item('SPECCODE').Value <> "UNAL" and  Tablers.Fields.Item('SPECCODE').Value <> "UNCO" and  Tablers.Fields.Item('SPECCODE').Value <> "UNCT" and  Tablers.Fields.Item('SPECCODE').Value <> "UNDU" and  Tablers.Fields.Item('SPECCODE').Value <> "UNGO" and  Tablers.Fields.Item('SPECCODE').Value <> "UNGU" and  Tablers.Fields.Item('SPECCODE').Value <> "UNJA" and  Tablers.Fields.Item('SPECCODE').Value <> "UNLA" and  Tablers.Fields.Item('SPECCODE').Value <> "UNLG" and  Tablers.Fields.Item('SPECCODE').Value <> "UNLO" and  Tablers.Fields.Item('SPECCODE').Value <> "UNLS" and  Tablers.Fields.Item('SPECCODE').Value <> "UNME" and  Tablers.Fields.Item('SPECCODE').Value <> "UNMU" and  Tablers.Fields.Item('SPECCODE').Value <> "UNPH" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSA" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSC" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSK" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSP" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSS" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSU" and  Tablers.Fields.Item('SPECCODE').Value <> "UNSW" and  Tablers.Fields.Item('SPECCODE').Value <> "UNTE" and  Tablers.Fields.Item('SPECCODE').Value <> "WFSP" and  Tablers.Fields.Item('SPECCODE').Value <> "WHIM" and  Tablers.Fields.Item('SPECCODE').Value <> "WISP" and  Tablers.Fields.Item('SPECCODE').Value <> "WTTR" and  Tablers.Fields.Item('SPECCODE').Value <> "WWSC"):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid SPECCODE"

            #Set and check VISIBLTY - should be done after LEGSTAGE is set, range=0-5
            if (Tablers.Fields.Item('VISIBLTY').Value is None) and (PrevVisiblty == 'Start'):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - First VISIBLTY is null. Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            elif Tablers.Fields.Item('VISIBLTY').Value is None and (PrevVisiblty <> 'Start'):
                Tablers.Fields.Item('VISIBLTY').Value = PrevVisiblty
            if Tablers.Fields.Item('VISIBLTY').Value < 0 or Tablers.Fields.Item('VISIBLTY').Value > 5:
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid VISIBLTY"

            #Set and check WX - should be done after LEGSTAGE is set, range=B,C,D,F,G,H,L,P,R,S,T,X
            if (Tablers.Fields.Item('WX').Value is None) and (PrevWX == 'Start'):
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - First WX is null. Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            elif Tablers.Fields.Item('WX').Value is None and (PrevWX <> 'Start'):
                Tablers.Fields.Item('WX').Value = PrevWX
            if Tablers.Fields.Item('WX').Value <> "B" and Tablers.Fields.Item('WX').Value <> "C" and Tablers.Fields.Item('WX').Value <> "D" and Tablers.Fields.Item('WX').Value <> "F" \
               and Tablers.Fields.Item('WX').Value <> "G" and Tablers.Fields.Item('WX').Value <> "H" and Tablers.Fields.Item('WX').Value <> "L" and Tablers.Fields.Item('WX').Value <> "P" \
               and Tablers.Fields.Item('WX').Value <> "R" and Tablers.Fields.Item('WX').Value <> "S" and Tablers.Fields.Item('WX').Value <> "T" and Tablers.Fields.Item('WX').Value <> "X":
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Invalid WX"


            #Calculate and check time elapsed (in decimal hours)
            CurrentHour = int(PMTimehour) + float(Timeminute)/60 + float(Timeseconds)/3600
            if PrevHour == 'Start':
                Tablers.Fields.Item('TIMELENGTH_dh').Value = 0  #set first record to 0
            else:
                Tablers.Fields.Item('TIMELENGTH_dh').Value = CurrentHour - PrevHour   #calculated as: time of current record - time of previous record
            if Tablers.Fields.Item('TIMELENGTH_dh').Value > maxtime: #Flag if > 15 minutes
                TimeDiff = round((Tablers.Fields.Item('TIMELENGTH_dh').Value*60), 1) #Convert time difference from decimal hours to minutes and round to 1 decimal place for output
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Time between events is " + str(TimeDiff) + " minutes. LEGSTAGE: " + str(Tablers.Fields.Item('LEGSTAGE').Value) + ". Notes: " + str(Tablers.Fields.Item('NOTES').Value)                

            #Calculate distance travelled (in meters)
            if PrevLat == 'Start':
                Tablers.Fields.Item('LENGTH_m').Value = 0 #set first record to 0
            else:
                Lat1 = Tablers.Fields.Item('LAT').Value
                Long1 = Tablers.Fields.Item('LONG').Value
                DistFunction = D.inv(PrevLong,PrevLat,Long1,Lat1, radians=False) #function uses geodetic distance module - calculates bearing and distance between 2 points
                Tablers.Fields.Item('LENGTH_m').Value = DistFunction[2]  #3rd component is the distance between 2 points

            #Calculate plane SPEED (in meters/second)
            if Tablers.Fields.Item('TIMELENGTH_dh').Value == 0: #set as null when no time has elapsed
                Tablers.Fields.Item('SPEED_ms').Value is None
            else:
                Tablers.Fields.Item('SPEED_ms').Value = (Tablers.Fields.Item('LENGTH_m').Value)/((Tablers.Fields.Item('TIMELENGTH_dh').Value)*3600) #equals distance travelled/time elapsed
            #Check plane speed is within limits
            if (Tablers.Fields.Item('SPEED_ms').Value is not None) and (Tablers.Fields.Item('SPEED_ms').Value < minspeed):
                SpeedKts = round(Tablers.Fields.Item('SPEED_ms').Value*1.94, 2) #Convert speed to knots and round to 2 decimal places for output
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Speed too low: " + str(SpeedKts) + " knots. LEGSTAGE: " + str(Tablers.Fields.Item('LEGSTAGE').Value) + ". Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            if (Tablers.Fields.Item('SPEED_ms').Value is not None) and (Tablers.Fields.Item('SPEED_ms').Value > maxspeed):
                SpeedKts = round(Tablers.Fields.Item('SPEED_ms').Value*1.94, 2)
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Speed too high: " + str(SpeedKts) + " knots. LEGSTAGE: " + str(Tablers.Fields.Item('LEGSTAGE').Value) + ". Notes: " + str(Tablers.Fields.Item('NOTES').Value)

            #Calculate and check altitude change (in meters/second)
            if Tablers.Fields.Item('ALT').Value is None:
                CurrentAlt = PrevAlt
            else:
                CurrentAlt = float(Tablers.Fields.Item('ALT').Value)
            if Tablers.Fields.Item('TIMELENGTH_dh').Value == 0: #set as null when no time has elapsed
                Tablers.Fields.Item('ALTCHANGE_ms').Value is None
            else:
                Tablers.Fields.Item('ALTCHANGE_ms').Value = (CurrentAlt - PrevAlt)/((Tablers.Fields.Item('TIMELENGTH_dh').Value)*3600) #equals difference in altitude/time elapsed
            if (Tablers.Fields.Item('ALTCHANGE_ms').Value is not None) and (Tablers.Fields.Item('ALTCHANGE_ms').Value < maxdescent):
                AltchangeFt = round(Tablers.Fields.Item('ALTCHANGE_ms').Value*196.9, 1) #Convert ALTCHANGE to ft/min and round to 1 decimal place for output
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Plane descent too fast: " + str(AltchangeFt) + " ft/min. LEGSTAGE: " + str(Tablers.Fields.Item('LEGSTAGE').Value) + ". Notes: " + str(Tablers.Fields.Item('NOTES').Value)
            if (Tablers.Fields.Item('ALTCHANGE_ms').Value is not None) and (Tablers.Fields.Item('ALTCHANGE_ms').Value > maxascent):
                AltchangeFt = round(Tablers.Fields.Item('ALTCHANGE_ms').Value*196.9, 1)
                print >> f, "EVENTNO: " + str(Tablers.Fields.Item('EVENTNO').Value) + " - Plane ascent too fast: " + str(AltchangeFt) + " ft/min. LEGSTAGE: " + str(Tablers.Fields.Item('LEGSTAGE').Value) + ". Notes: " + str(Tablers.Fields.Item('NOTES').Value)
                
            #Set current values as previous values before going to top of loop
            PrevFileID = Tablers.Fields.Item('FILEID').Value
            PrevEVENTNO = Tablers.Fields.Item('EVENTNO').Value
            PrevTime = Tablers.Fields.Item('TIME').Value
            if Tablers.Fields.Item('SIGHTNO').Value == 0:
                PrevSIGHTNO = PrevSIGHTNO
            elif Tablers.Fields.Item('SIGHTNO').Value <> 0:
                PrevSIGHTNO = Tablers.Fields.Item('SIGHTNO').Value
            PrevBeaufort = Tablers.Fields.Item('BEAUFORT').Value
            PrevCloud = Tablers.Fields.Item('CLOUD').Value
            PrevLegtype = Tablers.Fields.Item('LEGTYPE').Value
            if Tablers.Fields.Item('LEGSTAGE').Value is None or Tablers.Fields.Item('LEGSTAGE').Value == 5 and Tablers.Fields.Item('LEGSTAGE').Value <> 1:
                PrevLegstage = 5
            else:
                PrevLegstage = Tablers.Fields.Item('LEGSTAGE').Value
            PrevGlareL = Tablers.Fields.Item('GLAREL').Value
            PrevGlareR = Tablers.Fields.Item('GLARER').Value
            PrevVisiblty = Tablers.Fields.Item('VISIBLTY').Value
            PrevWX = Tablers.Fields.Item('WX').Value
            PrevHour = CurrentHour
            PrevLong = Tablers.Fields.Item('LONG').Value
            PrevLat = Tablers.Fields.Item('LAT').Value
            if Tablers.Fields.Item('ALT').Value is None:
                PrevAlt = CurrentAlt
            else:
                PrevAlt = float(Tablers.Fields.Item('ALT').Value)
            Tablers.MoveNext() #move to next record in table
            if Tablers.EOF and PrevLegstage <> 5: #Flag if last record not Legstage 5
                print >> f, "EVENTNO: " + str(PrevEVENTNO) + " - Last Legstage not 5"
            j = j +1


    SQLrs.MoveNext() #Move to next record (next table in list)
    i = i + 1
SQLconn.Close() #Close SQL connection

CPU_time2 = time.time()
CPU_time = (CPU_time2 - CPU_time1)/3600 #Calculate total processing time
print >> f, "Total processing time: %5.2f hours" % CPU_time

f.close() #Close the output text file

f = open(output, 'r') #Re-open the output text file for reading
for line in f:
    print line,  #print all lines in the output file to the Interactive Window
f.close() #closes the file

