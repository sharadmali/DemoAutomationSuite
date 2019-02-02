
	'==========================================================================================
	' Purpose		: Function to create snap shots
	' Author		: Cognizant Tecnology Solutions
	' Created on   	        : 
	' Last Updated	        : 
	' Reviewer		:
	'==========================================================================================



Class Screen_Shots
	Function Snap_Shots(ByVal Folder_Path,ByVal File_Name,ByRef objEnvironmentVariables)
		present_time=now
		present_time=Replace(present_time,"/","-")
		present_time=Replace(present_time," ","_")
		present_time=Replace(present_time,":","_")
		strpath=Folder_Path
		scr_path=Replace(strpath&"\"&File_Name&"_"&present_time&".png"," ","_")
		FileLen = Len(scr_path)
		If FileLen >= 259 then
			File_Name = Left(File_Name,Len(File_Name) - ( FileLen - 259))
			scr_path=Replace(strpath&"\"&File_Name&"_"&present_time&".png"," ","_")
		End If
		Browser("CreationTime:=0").CaptureBitmap scr_path
		objEnvironmentVariables.ScreenShotPath=scr_path
	End Function
End Class