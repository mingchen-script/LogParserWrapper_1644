# -LogParserWrapper_1644
# Readme..
	# This script will generate LDAP IP/Filters/TimeWaits summary Excel pivot tables from Directory Service's event 1644 EventLogs using LogParser and Excel via COM objects in 2 steps.
    #    1. Script calls LogParser to scans all event 1644 evtx in input directory, exact event data from event 1644, export to CSV.
    #    2. Script calls into Excel to import resulting CSV, create pivot tables for common ldap workload analysis. Delete CSV afterward.
	#
	# LogParserWrapper_1644.ps1 v0.6 11/16 (added DataBar)
	#		Steps: 
	#   	1. Install LogParser 2.2 from https://www.microsoft.com/en-us/download/details.aspx?id=24659
	#     	Note: More about LogParser2.2 https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-xp/bb878032(v=technet.10)?redirectedfrom=MSDN
	#   	2. Copy Directory Service EVTX from target DC(s) to same directory as this script.
	#     		Tip: When copying Directory Service EVTX, filter on event 1644 to reduce EVTX size for quicker transfer. 
	#					Note: Script will process all *.EVTX in script directory when run.
	#   	3. Run script
