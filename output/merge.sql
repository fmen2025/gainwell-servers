 MERGE  	into sts.server AS TR    
   USING    ( SELECT distinct LTRIM(RTRIM( [ita].[fnSnowServer](Servername,'Servername'))) AS Servername
			  ,Servername As OldServername
			  ,serverid
			  ,changeordernumber
			  ,ipaddress
			 ,operationalstatus
			 ,System
			 ,Process --rajitha added this column
			  ,Environment
			 ,location 
			  ,Sysclassname
			 ,Classification
			 ,os
			 ,osdomain
			 ,company
			 ,ram
			 ,dsnDomain
			 ,OSversion
			 ,supportedby
			 ,managedby
			 ,rn
			  ,Sysupdatedon 
			  ,syscreatedon
    FROM     (
	        -----retrive server information which have maximim Sysupdatedon
             SELECT		i.*, ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM( [ita].[fnSnowServer](i.Servername,'Servername')))
		     ORDER BY   Sysupdatedon  DESC,LTRIM(RTRIM( [ita].[fnSnowServer](i.Servername,'Servername')))) AS Rn    			
		     FROM       ITA.[Snowserver] i
				 --right join sts.server_demodelete d
				 --on d.id=i.serverid
		     where  	 lTRIM(RTRIM( [ita].[fnSnowServer] (Servername,'Servername'))) !='' 
			 or			 lTRIM(RTRIM( [ita].[fnSnowServer] (Servername,'Servername'))) is not null
					   --and 		servername= 'ah57c00wa04105'
					 --ltrim(rtrim(servername)) like 'ah57sdcub02001%'
             ) r
    WHERE    Rn = 1 
			 )  AS t 
			 On T.Servername   = [ita].[fnSnowServer](LTRIM(RTRIM(TR.Servername)),'Servername') 
			  ----SR.changeordernumber        = t.changeordernumber   

   
    WHEN   MATCHED --and  t.servername=tr.servername
    -- Server name change order number, ENV, Operation status, and classification
    THEN	  UPDATE 
    Set       tr.IPAddress= COALESCE(LTRIM(RTRIM(t.ipaddress)) ,tr.IPAddress) ---the reason behind only selected column need to update??
	          ,tr.ChangeOrderNumber= COALESCE(LTRIM(RTRIM(t.ChangeorderNumber)) ,tr.ChangeOrderNumber)
			  ,tr.operationalstatus= COALESCE(LTRIM(RTRIM(t.operationalstatus)) ,tr.operationalstatus)
			  ,tr.system= COALESCE(LTRIM(RTRIM(t.Process)) ,tr.system)
			  ,tr.primarysubsystem= COALESCE(LTRIM(RTRIM(t.environment)) ,tr.primarysubsystem)
			  ,tr.ENV= COALESCE(LTRIM(RTRIM(t.classification)) ,tr.ENV)
			  ,tr.Operatingsystem= COALESCE(LTRIM(RTRIM(t.os)) ,tr.Operatingsystem)
			  ,tr.Businessowner= COALESCE(LTRIM(RTRIM(t.managedby)) ,tr.Businessowner)
			  ,tr.ServerDescription = COALESCE(LTRIM(RTRIM(t.sysclassname)) ,tr.ServerDescription)
	          ,tr.DataCenter= COALESCE(LTRIM(RTRIM(t.location)) ,tr.DataCenter)
		     , tr.AssetOwner= COALESCE(LTRIM(RTRIM(t.osdomain)) ,tr.AssetOwner)
    		,  tr.Agency = COALESCE(LTRIM(RTRIM(t.company)) ,tr.Agency )
			  ,tr.memory= COALESCE(LTRIM(RTRIM(t.ram)) ,tr.memory)
	          ,tr.LocationExtension= COALESCE(LTRIM(RTRIM(t.dsnDomain)) ,tr.LocationExtension)
			  ,tr.OSversion= COALESCE(LTRIM(RTRIM(t.OSversion)) ,tr.OSversion)
              ,tr.TechnicalOwner= COALESCE(LTRIM(RTRIM(t.supportedby)) ,tr.TechnicalOwner)
	    ---if its source is null 
               ,tr.Source = COALESCE( tr.Source,'SNOWServer')
			   --if it is matched update the primary source'SnowServer' 
			   ,tr.primarySource= 'SNOWServer'
	           ,tr.BillingSource=COALESCE(tr.BillingSource, 'SNOWServer') 
			   ,tr.BillingType='On-Prem' 
			   ,tr.SnowServerFlag=1
			    ,tr.Decom=0
			  ;
			
