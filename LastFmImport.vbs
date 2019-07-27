Option Explicit
'==========================================================================
'
' MediaMonkey Script
'
' SCRIPTNAME: Last.fm Playcount Import
' DEVELOPMENT STARTED: 2009.02.17
  Dim Version : Version = "1.11"

' DESCRIPTION: Imports play counts from last.fm to update playcounts in MM
' FORUM THREAD: http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=15663&start=15#p191962
' 
' INSTALL: Copy to Scripts directory and add the following to Scripts.ini 
'          Don't forget to remove comments (') and set the order appropriately
'
'
' [LastFmImport]
' FileName=LastFmImport.vbs
' ProcName=LastFmImport
' Order=7
' DisplayName=Last FM Playcount Importer
' Description=Update missing playcounts from Last.fm
' Language=VBScript
' ScriptType=0 
'
' Changes: 1.11
' - Fix: More infalid xml characters checked
' - More graceful exits when errors occur
'
' Changes: 1.10
' - Fix: More invalid xml characters checked
'
' Changes: 1.9
' - Fix: Last.FM usernames not parsing correctly if containing special chars
'
' Changes: 1.8
' - Fix: Invalid ASCII characters stripped (hopefully - let me know if you find more!)
'	Thanks to SinDenial and AndréVonDrei for testing!
' - More graceful error messages (for some, let me know if you get anything cryptic)
' - Check for invalid characters when writing update file - some seem to cause errors
'	when the actual update went fine - needs improvment
'
' Changes: 1.7
' - Fix: Invalid apostrophes stripped, sadly this will make things less accurate
'	but will reduce error messages for the moment
'
' Changes: 1.6
' - Fix: No longer case sensitive for track names
' - Fix: Error messages on timeouts are more helpful
'
' Changes: 1.5
' Better logging - by default a log file will be created listing tracks updated
'	this file is LastFmImport.vbs.Updated.txt located in the scripts folder and is 
'	tab delimited.
' Better status bar messages as well I would like to think
'
' Changes: 1.4
' - HUGE speedup - no more .updateall() rather only update the track that needs it with
'   updateDB()
'
' Changes: 1.3
' - Database lookup optimizations (I hope!)
' - Added a few more stats to the process
' - Cleaned logging so enabling should only print essential stats about updated files
'
' Changes: 1.2
' - Status Bars!
' - Code Tidy Up
'
' Changes: 1.1
' - Abstracted username
' - Pretty error messages
'
'ToDo:
'* Smarter checking of files to update
'* Update LastPlayed time as well (if none exists)
'* Fix the update file writing to account for strange characters

Const ForReading = 1, ForWriting = 2, ForAppending = 8, Logging = True, Timeout = 25


Sub LastFMImport
	' Define variables
	Dim TrackChartXML, ChartListXML, DStart, DEnd, ArtistsL, TracksL
	' Stats variables
	Dim Plays, Matches, Counter, Updated, Tracks, Artists, ArtistMatches, LastMatch
	' Update logfile variables
	Dim fso, updatef


  
	' Status Bar
	Dim StatusBar
	Set StatusBar = SDB.Progress
  
	StatusBar.Text = "Getting UserName"

	dim uname
	uname=InputBox("Enter your Last.fm username:")

	Set ArtistsL = CreateObject("Scripting.Dictionary")

	StatusBar.Text = "Loading Weekly Charts List"
	Set ChartListXML = LoadXML(uname, "ChartList","","")
	SDB.ProcessMessages




	If Not (ChartListXML Is Nothing) Then
		If Not ChartListXML.getElementsByTagName("lfm").item(0).getAttribute("status") = "ok" Then
			MsgBox "Error" & VbCrLf & ChartListXML.getElementsByTagName("lfm").item(0).getElementsByTagName("error").item(0).text
			Exit Sub
		End If
		'logme " ChartListXML appears to be OK, proceeding with loading each weeks data"
		StatusBar.Text = "Loading Weekly Charts List -> OK"
		Dim Elem


		Plays = 0
		Counter = 0
		StatusBar.MaxValue = ChartListXML.getElementsByTagName("lfm").item(0).getElementsByTagName("weeklychartlist").item(0).getElementsByTagName("chart").length

		For Each Elem in ChartListXML.getElementsByTagName("lfm").item(0).getElementsByTagName("weeklychartlist").item(0).getElementsByTagName("chart")
		
			Counter = Counter + 1
			StatusBar.Text = "Loading Weekly Chart " & Counter & " of " & StatusBar.MaxValue
			StatusBar.Increase
			If StatusBar.Terminate Then
			  Exit Sub
			End If
			DStart = Elem.getAttribute("from")
			DEnd = Elem.getAttribute("to")


			logme " Attributes: " & DStart & " " & DEnd
			Set TrackChartXML = LoadXML(uname, "TrackChart",DStart,DEnd)
			SDB.ProcessMessages



			If NOT (TrackChartXML Is Nothing) Then
				If Not TrackChartXML.getElementsByTagName("lfm").item(0).getAttribute("status") = "ok" Then
					MsgBox "Error" & VbCrLf &  TrackListXML.getElementsByTagName("lfm").item(0).getElementsByTagName("error").item(0).text
					Exit Sub
				End If
				'logme "TrackChartXML appears to be OK, proceeding"
				Dim Ele, TrackTitle, ArtistName, PlayCount

			
				For Each Ele in TrackChartXML.GetElementsByTagName("lfm").item(0).GetElementsByTagName("track")

					TrackTitle = LCase(Ele.ChildNodes(1).Text)
					ArtistName = Ele.ChildNodes(0).ChildNodes(0).Text
					PlayCount = CInt(Ele.ChildNodes(3).Text)

					Plays = Plays + PlayCount

					'logme " < Searching for:> " &   ArtistName & " - " & TrackTitle & " = " & PlayCount & " Plays"

					If ArtistsL.Exists(ArtistName) Then
						If ArtistsL.Item(ArtistName).Exists(TrackTitle) Then
							ArtistsL.Item(ArtistName).Item(TrackTitle) = ArtistsL.Item(ArtistName).Item(TrackTitle) + PlayCount
						Else
							ArtistsL.Item(ArtistName).Add TrackTitle,PlayCount
						End If
					Else
						Dim temp
						Set temp = CreateObject("Scripting.Dictionary")
						temp.Add TrackTitle,PlayCount
						ArtistsL.Add ArtistName, temp

					End If
					

					SDB.ProcessMessages
				 

					
				Next
			Else
				Exit Sub
			End If
			SDB.ProcessMessages

		Next
		SDB.ProcessMessages

	Else
		logme "TracksListXML did not appear to load.. check loadxml() or network connection"
		msgbox ("Failed to get XML from LoadXML()")
	End If
	SDB.ProcessMessages


	ArtistMatches = 0
	Matches = 0
	Counter = 0
	StatusBar.Value = 0
	Updated = 0
	Tracks = 0
	Artists = ArtistsL.Count
	StatusBar.MaxValue = ArtistsL.Count
	StatusBar.Text = "Checking Database for Matches..."
	LastMatch = ""

	

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set updatef = fso.OpenTextFile(Script.ScriptPath&".Updated.txt",ForWriting,True)

	updatef.WriteLine "Artist" & VBTab & "Track" & VBTab & "New Plays" & VBTab & "Old Plays"
	For Each ArtistName In ArtistsL.Keys
		Dim list, ArtistTrackList
		SDB.ProcessMessages
		StatusBar.Increase
		StatusBar.Text = "Checking Database for Matches -> Updated: "  & Updated & "/" & Tracks & " Tracks " & LastMatch
		SDB.ProcessMessages
		'logme "Checking Database for Matches -> "  & StatusBar.Value & "/" & StatusBar.MaxValue & " -> " & ArtistName
		If StatusBar.Terminate Then
			Exit For
		End If

		SDB.ProcessMessages
		Set ArtistTrackList = ArtistsL.Item(ArtistName)
		SDB.ProcessMessages
		Tracks = Tracks + ArtistTrackList.Count
		SDB.ProcessMessages

		'Get all tracks in database by this artist
		Set list = QueryLibrary (ArtistName)
		SDB.ProcessMessages

		If list.Count > 0 Then
			Dim x
			SDB.ProcessMessages
			ArtistMatches = ArtistMatches + 1

			For x = 0 To list.Count-1
				Dim Item				
				SDB.ProcessMessages
				'logme "Loading next track by artist"
				Set Item = list.Item(x)
				SDB.ProcessMessages
				StatusBar.Text = "Checking Database for Matches -> Updated: "  & Updated & "/" & Tracks & " Tracks " & LastMatch
				'logme "Checking Database for Matches -> "  & StatusBar.Value & "/" & StatusBar.MaxValue & " -> " & ArtistName & " - " & list.Item(x).Title
				SDB.ProcessMessages
				If StatusBar.Terminate Then
					Exit For
				End If

				' Check if this track was on last.fm

				If ArtistTrackList.Exists(LCase(Item.Title)) Then
					SDB.ProcessMessages
					PlayCount = ArtistTrackList.Item(LCase(Item.Title))
					SDB.ProcessMessages

					Matches = Matches + 1

					logme " === Found: " & ArtistName & " - " & Item.Title & " PlayCount = " & PlayCount
					logme " === Previous plays: " & Item.PlayCounter
					logme " === LastPlayed: " & Item.LastPlayed

					If Item.PlayCounter < PlayCount Then 'Increase play count 
						LastMatch = " - Updating: " & ArtistName & " - " & Item.Title
						Updated = Updated + 1

						StatusBar.Text = "Checking Database for Matches -> Updated: "  & Updated & "/" & Tracks & " Tracks " & LastMatch
						
						'Unfortunatly, some files can be updated with wierd tags, yet cause errors when writing status file
						On Error Resume Next
						updatef.WriteLine ArtistName & VBTab & Item.Title & VBTab & PlayCount & VBTab & list.Item(x).PlayCounter
						If Err.Number <> 0 Then 
							numErr =Err.Number
							aboutErr = Err.description
							MsgBox "An Error has occured! Error number " & numerr & " of the type '" & abouterr & "'." & VbCrLf &_
								"Current artist was updated, but cannot be written to logfile."
							Err.Clear
						End If
						On Error Goto 0
						logme ArtistName & VBTab & Item.Title & VBTab & PlayCount & VBTab & list.Item(x).PlayCounter
						
						SDB.ProcessMessages
						list.Item(x).PlayCounter = PlayCount
						SDB.ProcessMessages
						
						
						logme " ==== Updating"
						SDB.ProcessMessages
						Item.UpdateDB()
					ElseIf Item.LastPlayed = 0.0 Then
							logme "Empty Last played"
					Else
						StatusBar.Text = "Checking Database for Matches -> Updated: "  & Updated & "/" & Tracks & " Tracks " & LastMatch
						'logme "Checking Database for Matches -> "  & StatusBar.Value & "/" & StatusBar.MaxValue &_								" -> SKIP: " & ArtistName & " - " & Item.Title
						'logme " ==== Skipping"
						SDB.ProcessMessages

					End If
					SDB.ProcessMessages
				Else
					SDB.ProcessMessages
					'logme " === Track not found: " & ArtistName & " - " & Item.Title
					
				End If
				SDB.ProcessMessages
			Next
			SDB.ProcessMessages
		Else
			SDB.ProcessMessages
			'logme "Artist does not exist: " & ArtistName
			SDB.ProcessMessages
		End If
		SDB.ProcessMessages
	Next
	Set fso = Nothing
	Set updatef = Nothing
	MsgBox  Plays & " Plays found on Last.fm consisting of " & Tracks & " tracks by " & Artists & " artists." & VbCrLf &_
		ArtistMatches & " of these artists were in the local database, along with " & Matches & " of their tracks." & VbCrLf &_
		"Tracks updated = " & Updated & VbCrLf & "The rest had a play count higher than last.fm already."
	
	SDB.ProcessMessages

End Sub

'**********************************************************


Function LoadXML(User,Mode,DFrom,DTo)
	'LoadXML accepts input string and mode, returns xmldoc of requested string and mode'
	'http://msdn2.microsoft.com/en-us/library/aa468547.aspx'
	logme ">> LoadXML: Begin with " & User & " & " & Mode
	Dim xmlDoc, xmlURL, StatusBar, LoadXMLBar, StartTimer, http, strippedText
	StartTimer = Timer

	Select Case Mode
		

		Case "ChartList"		'User Weekly Tracks Chart List
			xmlURL = "http://ws.audioscrobbler.com/2.0/?method=user.getWeeklyChartList&user=" &_
				fixurl(user) & "&api_key=daadfc9c6e9b2c549527ccef4af19adb"
		Case "TrackChart"		'User Weekly Tracks Chart
			xmlURL = "http://ws.audioscrobbler.com/2.0/?method=user.getweeklytrackchart&user=" &_
				fixurl(user) & "&api_key=daadfc9c6e9b2c549527ccef4af19adb&from=" & fixurl(dfrom) &_
				"&to=" & fixurl(dto)
		

	Case Else
		msgbox("Invalid MODE was passed to LoadXML(Input, Mode)")
		Exit Function
	End Select
	
	logme ">> URL: " & xmlURL

	Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
	Set http = CreateObject("Microsoft.XmlHttp")
	
	http.open "GET",xmlURL,True
	http.send ""
	

	StartTimer = Timer
	'Wait for up to 3 seconds if we've not gotten the data yet
	  Do While http.readyState <> 4 And Int(Timer-StartTimer) < Timeout
		SDB.ProcessMessages
		SDB.Tools.Sleep 100
		SDB.ProcessMessages
	  Loop

	  If (http.readyState <> 4) Then
		MsgBox ("HTTP request timed out. No tracks updated")
		Set LoadXML = Nothing
		Exit Function
	End If

	strippedText = stripInvalid(http.responseText)
	'MsgBox "Post Text: " & strippedText

	xmlDoc.async = True 
	xmlDoc.LoadXML(strippedText)

	StartTimer = Timer
	'Wait for up to 3 seconds if we've not gotten the data yet
	  Do While xmlDoc.readyState <> 4 And Int(Timer-StartTimer) < Timeout
		SDB.ProcessMessages
		SDB.Tools.Sleep 100
		SDB.ProcessMessages
	  Loop

	If (xmlDoc.parseError.errorCode <> 0) Then
		Dim myErr
		Set myErr = xmlDoc.parseError
		MsgBox("You have an error: " & myErr.reason)
		Set LoadXML = Nothing
	Else
		Dim currNode
		Set currNode = xmlDoc.documentElement.childNodes.Item(0)
	End If

	'logme " xmlDoc.Load: Waiting for Last.FM to return " & Mode & " of " & User
	SDB.ProcessMessages

	StartTimer = Timer
	Do While xmlDoc.readyState <> 4 And Int(Timer-StartTimer) < Timeout
		SDB.ProcessMessages
		SDB.Tools.Sleep 100
		SDB.ProcessMessages
	Loop



	'logme " xmlDoc: returned from loop in: " & (Timer - StartTimer)

	If xmlDoc.readyState = 4 and xmlDoc.parseError.errorCode = 0 Then 'all ok
		Set LoadXML = xmlDoc
		'msgbox("Last.FM query took: " & (timer-starttimer))
	Else
		'logme "Last.FM Query Failed @ " & Int(Timer-StartTimer) &	"ReadyState: " & xmlDoc.ReadyState & " URL: " & xmlURL
		msgbox("Last.FM Query Failed")
		Set LoadXML = Nothing 
	End if

	'logme "<< LoadXML: Finished in --> " & Int(Timer-StartTimer)

End Function

'******************************************************************
'**************** Library Query  **********************************
'******************************************************************

Function AddFilter()
	'logme " AddFilter(): Begin"
	'Add currently active filter to query if any'

	Dim GetFilter : GetFilter = SDB.Database.ActiveFilterQuery  
	If GetFilter <> "" Then
		AddFilter = " AND " & GetFilter
	End If
	
	'   'logme " AddFilter(): exit with :> " & AddFilter
	End Function

Function QueryLibrary(qArtist) 'input artist, title... output songlist'
	'logme "QueryLibrary: begin with " & qArtist
	Dim Iter, Iter2, Qry, Qry2, StatusBar, tmpSongList
	Set tmpSongList = SDB.NewSonglist

	Qry = "SELECT Artists.ID, Artists.Artist FROM Artists WHERE Artists.Artist LIKE '%" &_
		CorrectST(qArtist) & "'"

	'   Thanks to Bex for the improved qrys
	Qry2 = 	"AND Songs.ID IN (SELECT IDSong FROM ArtistsSongs, Artists  WHERE "&_
			"IDArtist=Artists.ID AND PersonType=1 AND " &_
			"Artist = '" & CorrectST(qArtist) & "') " & AddFilter
	'        "UpperW(TRIM(Artist)) = UpperW('" & CorrectST(qArtist) & "'))" & " " & Order
	'   'logme " QRY2 :> " & Qry2
	SDB.Database.Commit
	SDB.Database.BeginTransaction
	Set Iter = SDB.Database.OpenSQL(Qry)
	SDB.ProcessMessages
	Do While NOT Iter.EOF
		SDB.ProcessMessages
		'check artist exist first'
		'logme " found artist :> " & Iter.StringByIndex(1)

		Set Iter2 = SDB.Database.QuerySongs(Qry2)
		SDB.Database.Commit
		SDB.Database.BeginTransaction
		SDB.ProcessMessages
		Do While Not Iter2.EOF
		  SDB.ProcessMessages
		  tmpSongList.Add (Iter2.Item)
		  'logme " -->> Added: --->>> " & Iter2.Item.ArtistName & " - " & Iter2.Item.Title
	'       'logme Qry2
	'           msgbox("pause")
		  SDB.ProcessMessages
		  Iter2.Next
		  SDB.ProcessMessages
		Loop
		SDB.Database.Commit

		Iter.Next
		SDB.ProcessMessages
	Loop
	SDB.Database.Commit
	Set Iter = Nothing
	Set Iter2 = Nothing
	Set QueryLibrary = tmpSongList
End Function


'******************************************************************
'**************** Auxillary Functions *****************************
'******************************************************************

Sub logme(msg)
	'by psyXonova'
	If Logging Then
		'MsgBox "Yes!"
		Dim fso, logf
		On Error Resume Next
		Set fso = CreateObject("Scripting.FileSystemObject")
		'msgbox("logging: " & msg)
		Set logf = fso.OpenTextFile(Script.ScriptPath&".log",ForAppending,True)
		logf.WriteLine Now() & ": " & msg
		Set fso = Nothing
		Set logf = Nothing
	End If
End Sub



Function CorrectSt(inString)
' 	'logme ">> CorrectSt() has started with parameters: " & inString
	CorrectSt = Replace(inString, "'", "''")
	CorrectSt = Replace(CorrectSt, """", """""")
' 	'logme "<< CorrectSt() will return: " & CorrectSt & " and exit"
End Function


Function fixurl(sRawURL)
	' Original psyxonova improved by trixmoto
	logme ">> fixurl() entered with: " & sRawURL
	Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz\/!&:."
	sRawURL = Replace(sRawURL,"+","%2B")

	If UCase(Right(sRawURL,6)) = " (THE)" Then
		sRawURL = "The "&Left(sRawURL,Len(sRawURL)-6)
	End If
	If UCase(Right(sRawURL,5)) = ", THE" Then
		sRawURL = "The "&Left(sRawURL,Len(sRawURL)-5)
	End If

	If Len(sRawURL) > 0 Then
		Dim i : i = 1
		Do While i < Len(sRawURL)+1
			Dim s : s = Mid(sRawURL,i,1)
			If InStr(1,sValidChars,s,0) = 0 Then
				Dim d : d = Asc(s)
				If d = 32 Or d > 2047 Then
					s = "+"
				Else
					If d < 128 Then
						s = Hex(d)
					Else
						s = DecToUtf(d)
					End If
					s = "%" & s
				End If
			Else
				Select Case s
					Case "&"
						s = "%2526"
					Case "/"
						s = "%252F"
					Case "\"
						s = "%5C"
					Case ":"
						s = "%3A"
				End Select
			End If
			fixurl = fixurl&s
			i = i + 1
		SDB.ProcessMessages
    Loop
	End If
	logme "<< fixurl will return with: " & fixurl
End Function

Function stripInvalid(str)
	Dim re, newStr, invalidChars
	Set re = new regexp

	newStr = str
	' Need to do this better....
	invalidChars = Chr(1) & Chr(5) & Chr(6) & Chr(7) & Chr(12) & Chr(15) & Chr(16) & Chr(17) & Chr(23) & Chr(25) & Chr(31)
	re.Pattern = "[" & invalidChars & "]"
	Do While re.Test(newStr) = True
		newStr = re.Replace(newStr,"")
		'logme "==============Invalid character on this one!!???"
	Loop



	'logme "New text: " & VbCrLf & newStr & VbCrLf & "============================"
	stripInvalid = newStr
End Function

'************************************************************'

' Thanks to trixmoto for this function
Sub Install()
	Dim inip : inip = SDB.ApplicationPath&"Scripts\Scripts.ini"
	Dim inif : Set inif = SDB.Tools.IniFileByPath(inip)
	If Not (inif Is Nothing) Then
		inif.StringValue("LastFmImport","Filename") = "LastFmImport.vbs"
		inif.StringValue("LastFmImport","Procname") = "LastFmImport"
		inif.StringValue("LastFmImport","Order") = "7"
		inif.StringValue("LastFmImport","DisplayName") = "Last FM PlayCount Importer"
		inif.StringValue("LastFmImport","Description") = "Update missing playcounts from Last.fm"
		inif.StringValue("LastFmImport","Language") = "VBScript"
		inif.StringValue("LastFmImport","ScriptType") = "0"
		SDB.RefreshScriptItems
	End If
End Sub
