# note : the left from the --> (arow) indicate the column names list on the source [Gainwell Servers_20260820] excel templet and the right side of the arow  is the columns i maped to ita.snowserver by checking sample data from the templet.


System name --> System   # confirmed by name-match; DB column currently 100% NULL, no value overlap to verify against
Name --> servername
IP Address --> IPaddress 
Environment --> Environment 
Vendor	--> vendor
Location --> location
Purpose	--> Purpose   # confirmed by name-match; DB column currently 100% NULL, no value overlap to verify against
Used for --> usedfor
Install Status --> installstatus
Operational status --> operationalstatus
Operating System --> os
OS Version --> OSversion   # confirmed: matching non-null values (e.g. '7.9', '10') for same ServerName
Support group --> Supportgroup 
Supported by --> Supportedby   # confirmed: matching non-null values for same ServerName
Host admin group --> Hostadmin   # confirmed: matching non-null values for same ServerName
Host admin primary --> Hostadminprimary
Assigned to	--> (excluded)   # unmapped 2026-08-28; user determined Primaryapplicationadmin was not the right target column, no value overlap to verify against
Managed By Group --> Managedbygroup 
Managed by --> ManagedBy
Owned by --> Ownedby 
Classification --> Classification
Class --> Sysclassname
Maintenance schedule --> Maintenanceschedule
Discovery source --> (excluded)   # confirmed by user 2026-08-25; no matching column exists in ita.snowserver, not included in comparison
Most recent discovery --> Lastdiscovered   # name-match; Mostfrequentuser is a different concept (last logged-in user, not discovery date)
Updated	--> Sysupdatedon
Updated by --> Sysupdatedby
First discovered --> FirstDiscovered
Updates --> (excluded)   # confirmed by user 2026-08-25; sample values were close-but-not-identical to Sysmodcount, rejected as unreliable, not included in comparison
