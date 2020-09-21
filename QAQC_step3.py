# Author: Tim Gowan, FWRI
# Date: Feb 2011
# Purpose: After raw aerial survey data tables have been QC'd and edited, there will be a checked table for each survey day per team (e.g. FWRI20091204Check) in the SQL database.
#         This script combines the data from each of these tables into a single, final table.
#
#***The final table from this script replicates the historical DBF table created by URI/Bob Kenney
# Requirements: checked data tables (which have the suffix "Check" in name of table) from all teams which exist in SQL database

# Import system modules
import sys, string, os, math
import win32com.client
import time, datetime

CPU_time1 = time.time()  # Start time (value to estimate total processing time)

##################   Declare  Variables #######################################
SQLDB = "QAQC_2"  #Name of SQL database
yyyy = "0910"       #refers to years in season
############################################################################################################

# Connect to the SQL database 
SQL_DSN = 'PROVIDER=sqloledb;DATA SOURCE=.\SQLEXPRESS;Initial Catalog=' + SQLDB + ';Integrated Security=SSPI;'    # Connection string
SQLconn = win32com.client.Dispatch(r'ADODB.Connection')
SQLconn.Open(SQL_DSN) # Connect to the SQL server
SQLconn.Execute("SELECT name INTO Check_tables FROM sys.Tables WHERE name LIKE '%Check%' ORDER BY name") #Create new table with list of all Check tables
############################### for test runs only
try:
    SQLconn.Execute("DROP TABLE Final" + yyyy)
    print 'Deleted existing Final' + str(yyyy)
except:
    print 'Final' + str(yyyy) + 'has yet to be created' #delete Final table if it already exists
#############################
#Create empty "Final" table with the following fields and data types
SQLconn.Execute("CREATE TABLE Final" + yyyy + " (FILEID varchar(8), DDSOURCE varchar(3), MONTH smallint, DAY smallint, YEAR smallint, EVENTNO smallint, TIME int, LATDEG smallint, LATMIN numeric(7,5), LONGDEG smallint, LONGMIN numeric(7,5), LEGTYPE smallint, LEGSTAGE smallint, ALT numeric(5,1), HEADING numeric(4,1), WX varchar(1), VISIBLTY numeric(3,1), BEAUFORT smallint, CLOUD smallint, GLAREL smallint, GLARER smallint, SIGHTNO smallint, SPECCODE varchar(4), IDREL smallint, NUMBER smallint, CONFIDNC smallint, NUMCALF smallint, ANHEAD smallint, PHOTOS smallint, BEHAV1 smallint, BEHAV2 smallint, BEHAV3 smallint, BEHAV4 smallint, BEHAV5 smallint, BEHAV6 smallint, BEHAV7 smallint, BEHAV8 smallint, BEHAV9 smallint, BEHAV10 smallint, BEHAV11 smallint, BEHAV12 smallint, BEHAV13 smallint, BEHAV14 smallint, BEHAV15 smallint)")

SQLrs = win32com.client.Dispatch(r'ADODB.Recordset')
SQLrs.Open('SELECT * FROM Check_tables ORDER BY name', SQLconn, 1, 3)   # Open list of Check tables as a recordset
Ntables = SQLrs.RecordCount
print 'There are ' + str(Ntables) + ' total Check tables.'

i = 0
SQLrs.MoveFirst() #Move to first record (first table in list)

# Loop through each Check table, passing over "Check_tables"
#####change to < Ntables in actual run to check all tables
while i < Ntables: #while there are still tables to append
    stringname = SQLrs.Fields.Item('name').Value
    suffix = "tables"
    if stringname.endswith(suffix): #pass over "Check_tables"
        pass
    else:
        # Begin appending tables
        Table = SQLrs.Fields.Item('name').Value     #Table is all data from a single survey day
        print "Adding table " + Table
        #Insert data from "Check" table into "Final" table
        SQLconn.Execute("INSERT INTO Final" + yyyy + " SELECT FILEID, DDSOURCE, MONTH, DAY, YEAR, EVENTNO, TIME, LATDEG, LATMIN, LONGDEG, LONGMIN, LEGTYPE, LEGSTAGE, ALT, HEADING, WX, VISIBLTY, BEAUFORT, CLOUD, GLAREL, GLARER, SIGHTNO, SPECCODE, IDREL, NUMBER, CONFIDNC, NUMCALF, ANHEAD, PHOTOS, BEHAV1, BEHAV2, BEHAV3, BEHAV4, BEHAV5, BEHAV6, BEHAV7, BEHAV8, BEHAV9, BEHAV10, BEHAV11, BEHAV12, BEHAV13, BEHAV14, BEHAV15 FROM " + Table + " ORDER BY EVENTNO")
        
    SQLrs.MoveNext() #Move to next record (next table in list)
    i = i + 1
SQLconn.Close() #Close SQL connection

CPU_time2 = time.time()
CPU_time = (CPU_time2 - CPU_time1)/3600 #Calculate total processing time
print 'Total processing time: %5.2f hours' % CPU_time

