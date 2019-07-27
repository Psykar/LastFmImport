Option Explicit
'==========================================================================
'
' MediaMonkey Script
'
' SCRIPTNAME: Last.fm Playcount Import
' DEVELOPMENT STARTED: 2009.02.17
  Dim Version : Version = "1.2"

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
' Changes: 1.2
' - Status Bars!
' - Code Tidy Up
'
' Changes: 1.1
' - Abstracted username
' - Pretty error messages
'
' TODO: 
' * Smarter checking of files to update

Const ForReading = 1, ForWriting = 2, ForAppending = 8, Logging = True, Timeout = 25


Sub LastFMImport
	' Define variables
	Dim TrackChartXML, ChartListXML, DStart, DEnd, ArtistsL, TracksL
	' Stats variables
	Dim Plays, Skipped, Matches, UniqueMatches, Counter

  
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

	If Not ChartListXML.getElementsByTagName("lfm").item(0).getAttribute("status") = "ok" Then
		MsgBox "Error" & VbCrLf & ChartListXML.getElementsByTagName("lfm").item(0).getElementsByTagName("error").item(0).text
		Exit Sub
	End If


	If Not (ChartListXML Is Nothing) Then
		logme " ChartListXML appears to be OK, proceeding with loading each weeks data"
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
			If Not TrackChartXML.getElementsByTagName("lfm").item(0).getAttribute("status") = "ok" Then
				MsgBox "Error" & VbCrLf &  TrackListXML.getElementsByTagName("lfm").item(0).getElementsByTagName("error").item(0).text
				Exit Sub
			End If


			If NOT (TrackChartXML Is Nothing) Then
				logme "TrackChartXML appears to be OK, proceeding"
				Dim Ele, TrackTitle, ArtistName, PlayCount

			
				For Each Ele in TrackChartXML.GetElementsByTagName("lfm").item(0).GetElementsByTagName("track")

					TrackTitle = Ele.ChildNodes(1).Text
					ArtistName = Ele.ChildNodes(0).ChildNodes(0).Text
					PlayCount = CInt(Ele.ChildNodes(3).Text)

					Plays = Plays + PlayCount

					logme " < Searching for:> " &   ArtistName & " - " & TrackTitle & " = " & PlayCount & " Plays"

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
				msgbox("did not get any matches from Chart tracks xml")
			End If

		Next

	Else
		logme "TracksListXML did not appear to load.. check loadxml() or network connection"
		msgbox ("Failed to get XML from LoadXML()")
	End If




	Skipped = 0
	Matches = 0
	UniqueMatches = 0
	Counter = 0
	StatusBar.Value = 0
	StatusBar.MaxValue = ArtistsL.Count
	StatusBar.Text = "Checking Database for Matches..."



	For Each ArtistName In ArtistsL.Keys

		StatusBar.Increase
		StatusBar.Text = "Checking Database for Artist Matches -> "  & StatusBar.Value & "/" & StatusBar.MaxValue & " -> " & ArtistName
		If StatusBar.Terminate Then
			Exit For
		End If
		For Each TrackTitle In ArtistsL.Item(ArtistName).Keys
			Dim list
			StatusBar.Text = "Checking Database for Artist Matches -> "  & StatusBar.Value & "/" & StatusBar.MaxValue & " -> " & ArtistName &_
				" - " & TrackTitle
			If StatusBar.Terminate Then
				Exit For
			End If
			PlayCount = ArtistsL.Item(ArtistName).Item(TrackTitle)

			Set list = QueryLibrary (ArtistName,TrackTitle)
			If list.Count > 0 Then

				Dim x

				Matches = Matches + list.Count
				UniqueMatches = UniqueMatches + 1

				logme " == Found:> " &   ArtistName & " - " & TrackTitle & " = " & PlayCount & " Plays"
				For x = 0 to list.Count-1
					logme " === Previous plays: " & list.Item(x).PlayCounter
					If list.Item(x).PlayCounter < PlayCount Then 'Increase play count 
							StatusBar.Text = "Checking Database for Artist Matches -> "  & StatusBar.Value & "/" & StatusBar.MaxValue & " -> MATCH: " & ArtistName &_
								" - " & TrackTitle
						list.Item(x).PlayCounter = PlayCount
						logme " ==== Updating"
					Else
							StatusBar.Text = "Checking Database for Artist Matches -> "  & StatusBar.Value & "/" & StatusBar.MaxValue & " -> SKIP: " & ArtistName &_
								" - " & TrackTitle
						Skipped = Skipped + 1
						logme " ==== Skipping"
					End If
					SDB.ProcessMessages
				Next

			End If
			list.UpdateAll
		Next
	Next

	MsgBox "Matches to local = " & Matches & VbCrLf & "UNIQUE Matches to local = " & UniqueMatches &_
		VbCrLf & "Plays found on Last.fm = " & Plays & VbCrLf & "Skipped as >= plays on last.fm = " &_
		Skipped & VbCrLf & "Tracks updated = " & Matches - Skipped
	
	SDB.ProcessMessages



End Sub

'**********************************************************


Function LoadXML(User,Mode,DFrom,DTo)
	'LoadXML accepts input string and mode, returns xmldoc of requested string and mode'
	'http://msdn2.microsoft.com/en-us/library/aa468547.aspx'
	logme ">> LoadXML: Begin with " & User & " & " & Mode
	Dim xmlDoc, xmlURL, StatusBar, LoadXMLBar, StartTimer, http
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

	Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
	Set http = CreateObject("Microsoft.XmlHttp")
	
	http.open "GET",xmlURL,False
	http.send ""
	

	xmlDoc.async = True 
	xmlDoc.LoadXML(http.responseText)

	If (xmlDoc.parseError.errorCode <> 0) Then
		Dim myErr
		Set myErr = xmlDoc.parseError
		MsgBox("You have error " & myErr.reason)
	Else
		Dim currNode
		Set currNode = xmlDoc.documentElement.childNodes.Item(0)
	End If

	logme " xmlDoc.Load: Waiting for Last.FM to return " & Mode & " of " & User
	SDB.ProcessMessages

	
	Do While xmlDoc.readyState <> 4 And Int(Timer-StartTimer) < Timeout
		SDB.ProcessMessages
		SDB.Tools.Sleep 100
		SDB.ProcessMessages
	Loop



	logme " xmlDoc: returned from loop in: " & (Timer - StartTimer)

	If xmlDoc.readyState = 4 Then 'all ok
		Set LoadXML = xmlDoc
		'msgbox("Last.FM query took: " & (timer-starttimer))
	Else
		logme "Last.FM Query Failed @ " & Int(Timer-StartTimer) &_
		"ReadyState: " & xmlDoc.ReadyState & " URL: " & xmlURL
		msgbox("Last.FM Timed Out @ " & Int(Timer-StartTimer))
		Set LoadXML = Nothing 
	End if

	logme "<< LoadXML: Finished in --> " & Int(Timer-StartTimer)

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
	
	'   logme " AddFilter(): exit with :> " & AddFilter
	End Function

	Function QueryLibrary(qArtist, qTitle) 'input artist, title... output songlist'
	logme "QueryLibrary: begin with " & qArtist & ", " & qTitle
	Dim Iter, Iter2, Qry, Qry2, StatusBar, tmpSongList
	Set tmpSongList = SDB.NewSonglist

	Qry = "SELECT Artists.ID, Artists.Artist FROM Artists WHERE Artists.Artist LIKE '%" &_
		CorrectST(qArtist) & "'"

	'   Thanks to Bex for the improved qrys
	Qry2 = "SongTitle = '" & CorrectST(qTitle) & "' "&_
			"AND Songs.ID IN (SELECT IDSong FROM ArtistsSongs, Artists  WHERE "&_
			"IDArtist=Artists.ID AND PersonType=1 AND " &_
			"Artist = '" & CorrectST(qArtist) & "') " & AddFilter
	'        "UpperW(TRIM(Artist)) = UpperW('" & CorrectST(qArtist) & "'))" & " " & Order
	'   logme " QRY2 :> " & Qry2
	SDB.Database.Commit
	SDB.Database.BeginTransaction
	Set Iter = SDB.Database.OpenSQL(Qry)
	SDB.ProcessMessages
	Do While NOT Iter.EOF
		SDB.ProcessMessages
		'check artist exist first'
		logme " found artist :> " & Iter.StringByIndex(1) & " for track " & qTitle

		Set Iter2 = SDB.Database.QuerySongs(Qry2)
		SDB.Database.Commit
		SDB.Database.BeginTransaction
		SDB.ProcessMessages
		Do While Not Iter2.EOF
		  SDB.ProcessMessages
		  tmpSongList.Add (Iter2.Item)
		  logme " -->> Added: --->>> " & Iter2.Item.ArtistName & " - " & Iter2.Item.Title
	'       logme Qry2
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
' 	logme ">> CorrectSt() has started with parameters: " & inString
	CorrectSt = Replace(inString, "'", "''")
	CorrectSt = Replace(CorrectSt, """", """""")
' 	logme "<< CorrectSt() will return: " & CorrectSt & " and exit"
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
						s = DecToHex(d)
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
