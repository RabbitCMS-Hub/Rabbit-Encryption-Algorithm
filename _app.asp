<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************
'**********************************************
' RES - Rabbit Encryption System
' Special Thanks: 
' 	RC4 Encryption Using ASP & VBScript By Mike Shaffer
' 	Encrypt and decrypt functions for classic ASP (by TFI)
' 	This class contains specially compiled and enhanced functions.
' 	https://github.com/badursun/Classic-ASP-Encryption-Decryption-With-Special-Key-Class/
'**********************************************

Class Rabbit_Encryption_Algorithm
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME, PLUGIN_AUTOLOAD
	Private g_KeyLen, g_KeyLocation, g_DefaultKey, cryptkey, SavePathLocation, fileSaveState, LastSavedFileName

	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		
		' Check Register
		'------------------------------
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE)
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "Rabbit_Encryption_Algorithm")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "554")
		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "0")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)

		' Plugin Settings
		'------------------------------

		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Plugin Admin Panel Extention
	'---------------------------------------------------------------
	Public sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "SHOW:SampleCode" Then
			Call PluginPage("Header")

			SpecialWords    = "Bu gizli bir kelimedir, gizli olarak kalması gerekmektedir!"
			SampleKey 		= "/[g-?#4Fd$-T/{\d3%%.@!specialLongKey$%&₺sW"
			Key  			= SampleKey
			MyKeyFileName 	= "mykey.txt"

			With Response 
				.Write "<div class=""row"">"
				.Write "	<div class=""col-lg-6 col-12>"
				.Write "		<h6>Code Syntax</h6>"
				.Write "	</div>"
				.Write "	<div class=""col-lg-6 col-12>"
				.Write "		<h6>Sample Output</h6>"
				.Write "		<style>td{width:33%; word-wrap:break-word;}</style>"
				.Write "		<table class=""tabel table-sm table-bordered"">"
				.Write "			<tr>"
				.Write "				<td>Sample String</td>"
				.Write "				<td></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> "& SpecialWords &"</small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Let Key</td>"
				.Write "				<td><code>Encryption.Key = """& SampleKey &"""</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> </small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Default Key</td>"
				.Write "				<td><code>Encryption.DefaultKey()</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> "& SampleKey &"</small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Get Key</td>"
				.Write "				<td><code>Encryption.Key()</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> </small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Encrpyted</td>"
				.Write "				<td><code>Encryption.Encrypt(SpecialWords)</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> "& Encrypt(SpecialWords) &"</small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Decrypted</td>"
				.Write "				<td><code>Encryption.Decrypt("""& Encrypt(SpecialWords) &""")</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> "& Decrypt( Encrypt(SpecialWords) ) &"</small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Key Save Path</td>"
				.Write "				<td><code>Encryption.KeySavePath()</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> "& KeySavePath() &"</small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Save Key to File</td>"
				.Write "				<td><code>Encryption.WriteKeyToFile("""& MyKeyFileName &""")</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> "& WriteKeyToFile(MyKeyFileName) &"</small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Readed key from File?</td>"
				.Write "				<td><code>Encryption.ReadKeyFromFile("""& MyKeyFileName &""")</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> "& ReadKeyFromFile(MyKeyFileName) &"</small></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td>Readed key from File? (Not Exist file)</td>"
				.Write "				<td><code>Encryption.ReadKeyFromFile(""NotExistFile.txt"")</code></td>"
				.Write "			</tr>"
				.Write "			<tr>"
				.Write "				<td colspan=""2""><small><strong>RESULT:</strong> "& ReadKeyFromFile("NotExistFile.txt") &"</small></td>"
				.Write "			</tr>"
				.Write "		</table>"
				.Write "	</div>"
				.Write "</div>"
			End With


			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			' .Write "<div class=""row"">"
			' .Write "    <div class=""col-lg-6 col-sm-12"">"
			' .Write 			QuickSettings("select", ""& PLUGIN_CODE &"_OPTION_1", "Buraya Title", "0#Seçenek 1|1#Seçenek 2|2#Seçenek 3", TO_DB)
			' .Write "    </div>"
			' .Write "    <div class=""col-lg-6 col-sm-12"">"
			' .Write 			QuickSettings("number", ""& PLUGIN_CODE &"_OPTION_2", "Buraya Title", "", TO_DB)
			' .Write "    </div>"
			' .Write "    <div class=""col-lg-12 col-sm-12"">"
			' .Write 			QuickSettings("tag", ""& PLUGIN_CODE &"_OPTION_3", "Buraya Title", "", TO_DB)
			' .Write "    </div>"
			' .Write "</div>"
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:SampleCode"" class=""btn btn-sm btn-primary"">"
			.Write "        	[DEV] Örnek Kullanım"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Class First Init
	'---------------------------------------------------------------
	Private Sub class_initialize()
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_CODE  			= "RABENCYRPTALGO"
    	PLUGIN_NAME 			= "Rabbit Encryption Algorithm"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/Rabbit-Encryption-Algorithm"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-memory"
    	PLUGIN_CREDITS 			= "RC4 Encryption Using ASP & VBScript By @Mike Shaffer ReDeveloped by @badursun Anthony Burak DURSUN"
    	PLUGIN_FOLDER_NAME 		= "Rabbit-Encryption-Algorithm"
    	PLUGIN_DB_NAME 			= ""
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_AUTOLOAD 		= True
    	PLUGIN_ROOT 			= PLUGIN_DIST_FOLDER_PATH(This)
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------

		fileSaveState 		= False
		LastSavedFileName 	= ""
		g_KeyLocation 		= PLUGIN_FILES_ROOT + "keyFiles/"
		g_KeyLen 			= 512
		g_DefaultKey 		= "GNQ?4i0-*\CldnU+[vrF1j1PcWeJfVv4QGBurFK6}*l[H1S:oY\v@U?i" &_
							  ",oD]f/n8oFk6NesH--^PJeCLdp+(t8SVe:ewY(wR9p-CzG<,Q/(U*.pX" &_ 
							  "Diz/KvnXP`BXnkgfeycb)1A4XKAa-2G}74Z8CqZ*A0P8E[S`6RfLwW+P" &_ 
							  "c}13U}_y0bfscJ<vkA[JC;0mEEuY4Q,([U*XRR}lYTE7A(O8KiF8>W/m" &_
							  "1D*YoAlkBK@`3A)trZsO5xv@5@MRRFkt\"

		cryptkey 			= g_DefaultKey
		SavePathLocation	= Server.Mappath( g_KeyLocation )

    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Register App
    	'-------------------------------------------------------------------------------------
    	class_register()

    	'-------------------------------------------------------------------------------------
    	' Hook Auto Load Plugin
    	'-------------------------------------------------------------------------------------
    	If PLUGIN_AUTOLOAD_AT("WEB") = True Then 

    	End If
	End Sub
	'---------------------------------------------------------------
	' Class First Init
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Class Terminate
	'---------------------------------------------------------------
	Private sub class_terminate()

	End Sub
	'---------------------------------------------------------------
	' Class Terminate
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property
	Public Property Get PluginAutoload() 	: PluginAutoload = PLUGIN_AUTOLOAD 			: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable, PluginAutoload)
	End Property
	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------



	'---------------------------------------------------------------
	' Get Default Key
	'---------------------------------------------------------------
	Public Property Get KeySavePath()
		KeySavePath = SavePathLocation
	End Property
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Get Default Key
	'---------------------------------------------------------------
	Public Property Get DefaultKey()
		DefaultKey = g_DefaultKey
	End Property
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Let cryptkey
	'---------------------------------------------------------------
	Public Property Let FilePath(vVal)
		If Len(vVal) < 1 Then 
			SavePathLocation = Server.Mappath( g_KeyLocation )
		Else 
			SavePathLocation = Server.Mappath( vVal )
		End If
	End Property
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Let cryptkey
	'---------------------------------------------------------------
	Public Property Let Key(vVal)
		If Len(vVal) < 1 Then 
			cryptkey = KeyGeN( g_KeyLen )
		Else 
			cryptkey = vVal
		End If
	End Property
	Public Property Get Key()
		Key = cryptkey
	End Property
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Let File Saved
	'---------------------------------------------------------------
	Public Property Get LastSavedFile()
		LastSavedFile = LastSavedFileName
	End Property
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Let File Saved
	'---------------------------------------------------------------
	Public Property Get FileSaved()
		FileSaved = fileSaveState
	End Property
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Write Key File Pyhsical Folder
	'---------------------------------------------------------------
	Public Function WriteKeyToFile(strFileName)
		LastSavedFileName = ""
		On Error Resume Next

		Dim PVFso, cPVFso
	    Set PVFso = Server.CreateObject("Scripting.FileSystemObject")
	        If PVFso.FolderExists(SavePathLocation) = False Then
	            Set cPVFso = CreateObject("Scripting.FileSystemObject" )
	                cPVFso.CreateFolder(SavePathLocation)
	            Set cPVFso = Nothing

	            Call PanelLog(""& PLUGIN_CODE &" Plugin için "& SavePathLocation &" klasörü oluşturuldu.", 0, ""& PLUGIN_CODE &"", 0)
	        End If
	    Set PVFso = Nothing

		Dim keyFile, fso
		set fso = Server.CreateObject("scripting.FileSystemObject") 
		set keyFile = fso.CreateTextFile(SavePathLocation&"\"&strFileName, true) 
			keyFile.WriteLine( cryptkey )
			keyFile.Close
		
		If Err <> 0 Then
			fileSaveState = False
		Else
			fileSaveState = True
			LastSavedFileName = strFileName
		End If

		WriteKeyToFile = fileSaveState

		set keyFile = Nothing
		set fso = Nothing
	End Function
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Read Key From File 
	'---------------------------------------------------------------
	Public Function ReadKeyFromFile(strFileName)
		Dim keyFile, Fso, f
		Set Fso = Server.CreateObject("Scripting.FileSystemObject") 
		
		tmpFileFullPath = SavePathLocation&"\"&strFileName

		If Not Fso.FileExists( tmpFileFullPath ) = True Then
			ReadKeyFromFile = "File Not Found"
		Else
			Set f = Fso.GetFile( tmpFileFullPath ) 
			Set ts = f.OpenAsTextStream(1, -2)

			Do While not ts.AtEndOfStream
				keyFile = keyFile & ts.ReadLine
			Loop 

			ReadKeyFromFile =  keyFile
		End If

		Set ts = Nothing 
		Set f = Nothing 
		Set Fso = Nothing
	End Function
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' URL Encode/Decode
	'---------------------------------------------------------------
	Private Function URLDecode4Encrypt(sConvert)
		Dim aSplit
		Dim sOutput
		Dim I
		If IsNull(sConvert) Then
			URLDecode4Encrypt = ""
			Exit Function
		End If
		
		'sOutput = REPLACE(sConvert, "+", " ") ' convert all pluses to spaces
		sOutput=sConvert
		aSplit = Split(sOutput, "%")
		If IsArray(aSplit) Then
			sOutput = aSplit(0)
			For I = 0 to UBound(aSplit) - 1
				sOutput = sOutput &  Chr("&H" & Left(aSplit(i + 1), 2)) & Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
			Next
		End If
		URLDecode4Encrypt = sOutput
	End Function
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' KeyGen generator
	'---------------------------------------------------------------
	Private Function KeyGeN(iKeyLength)
		Dim k, iCount, strMyKey
		lowerbound = 35 
		upperbound = 96
		Randomize
		for i = 1 to iKeyLength
			s = 255
			k = Int(((upperbound - lowerbound) + 1) * Rnd + lowerbound)
			strMyKey =  strMyKey & Chr(k) & ""
		next
		KeyGeN = strMyKey
	End Function
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Encrypt
	'---------------------------------------------------------------
	Public function Encrypt(inputstr)
		Dim i,x
		outputstr=""
		cc=0
		for i=1 to len(inputstr)
			x=asc(mid(inputstr,i,1))
			x=x-48
			if x<0 then x=x+255
			x=x+asc(mid(cryptkey,cc+1,1))
			if x>255 then x=x-255
			outputstr=outputstr&chr(x)
			cc=(cc+1) mod len(cryptkey)
		next
		Encrypt = server.urlencode(replace(outputstr,"%","%25"))
	end function
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Decrypt
	'---------------------------------------------------------------
	Public function Decrypt(byval inputstr)
		Dim i,x
		inputstr=URLDecode4Encrypt(inputstr)
		outputstr=""
		cc=0
		for i=1 to len(inputstr)
			x=asc(mid(inputstr,i,1))
			x=x-asc(mid(cryptkey,cc+1,1))
			if x<0 then x=x+255
			x=x+48
			if x>255 then x=x-255
			outputstr=outputstr&chr(x)
			cc=(cc+1) mod len(cryptkey)
		next
		Decrypt = outputstr
	end function
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' HMAC-MD5 Ecrypt (Message, Key)
	'---------------------------------------------------------------
	Public Function HMACMD5(Message, Key)
	    On Error Resume Next

	    Dim PlainText
	    	PlainText = Message

	    With CreateObject("ADODB.Stream")
	        .Open
	        .CharSet = "Windows-1252"
	        .WriteText PlainText
	        .Position = 0
	        .CharSet = "UTF-8"
	        PlainText = .ReadText
	        .Close
	    End With

	    Set UTF8Encoding = CreateObject("System.Text.UTF8Encoding")
	    Dim PlainTextToBytes, BytesToHashedBytes, HashedBytesToHex

	    PlainTextToBytes = UTF8Encoding.GetBytes_4(PlainText)

	    Set Cryptography = CreateObject("System.Security.Cryptography.HMACMD5")
	    	Cryptography.Initialize()
	    	Cryptography.Key = UTF8Encoding.GetBytes_4( Key )

	    BytesToHashedBytes = Cryptography.ComputeHash_2((PlainTextToBytes))

	    For x = 1 To LenB(BytesToHashedBytes)
	        HashedBytesToHex = HashedBytesToHex & Right("0" & Hex(AscB(MidB(BytesToHashedBytes, x, 1))), 2)
	    Next

		If Err.Number <> 0 Then 
			Response.Write(Err.Description) 
		Else 
			HMACMD5 = LCase(HashedBytesToHex)
		End If

	    Set Cryptography = Nothing
	    Set UTF8Encoding = Nothing
	    On Error GoTo 0
	End Function
	'---------------------------------------------------------------
	' END
	'---------------------------------------------------------------
End Class



' Set Encryption = New RabbitEncryptionSystem
'   SpecialWords    = "Bu gizli bir kelimedir, gizli olarak kalması gerekmektedir!"

'   ' If you want create random Key uncomment end define empty string
'   ' Or Define your Key. More Length, More Strength (512 Length)
'   Encryption.Key  = "/[g-?#4Fd$-T/{\d3%%.@!specialLongKey$%&₺sW"


'   Response.Write "<strong>Default Key:</strong> " & Encryption.DefaultKey()
'   Response.Write "<hr>"

'   Response.Write "<strong>Key:</strong> " & Encryption.Key()
'   Response.Write "<hr>"
'   Response.Write "<strong>Keyword:</strong> " & SpecialWords
'   Response.Write "<hr>"
'   Response.Write "<strong>Encrypted:</strong> " & Encryption.Encrypt(SpecialWords)
'   Response.Write "<hr>"
'   Response.Write "<strong>Decrypted:</strong> " & Encryption.Decrypt( Encryption.Encrypt(SpecialWords) )
'   Response.Write "<hr>"

'   'Encryption.FilePath = "../cache/"
'   Response.Write "Key Save Path: " & Encryption.KeySavePath()
'   Response.Write "<hr>"

'   ' Save File To Path
'   MyKeyFileName = "mykey.txt"
'   Encryption.WriteKeyToFile(MyKeyFileName)
'   Response.Write "<strong>File Saved:</strong> " & Encryption.FileSaved()
'   Response.Write "<hr>"
'   Response.Write "<strong>Last Saved File Name:</strong> " & Encryption.LastSavedFile()
'   Response.Write "<hr>"

'   ' Read File For Key
'   Response.Write "<strong>Readed Key File:</strong> " & MyKeyFileName
'   Response.Write "<hr>"
'   Response.Write "<strong>Readed Key Value:</strong> " & Encryption.ReadKeyFromFile(MyKeyFileName)
'   Response.Write "<hr>"
'   Response.Write "<strong>Not Exist Readed Key File:</strong> " & Encryption.ReadKeyFromFile("NotExistFile.txt")
'   Response.Write "<hr>"

' Set Encryption = Nothing
%>
