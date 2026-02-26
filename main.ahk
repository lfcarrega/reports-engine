#Requires AutoHotkey v2.0
#Include adosql.ahk

; SQL Server 2022 LocalDB:
; https://download.microsoft.com/download/3/8/d/38de7036-2433-4207-8eae-06e247e17b25/SqlLocalDB.msi

FileInstall("lock-24.png", "lock-24.png", 1)
FileInstall("user-24.png", "user-24.png", 1)
FileInstall("file-analytics-64.png", "file-analytics-64.png", 1)
FileInstall("main.ico", "main.ico", 1)

If ! DirExist("bin\win\amd64")
	DirCreate("bin\win\amd64")
FileInstall("bin\win\amd64\caddy.exe", "bin\win\amd64\caddy.exe", 1)

If ! DirExist("srv\js")
	DirCreate("srv\js")
FileInstall("srv\js\jquery_wrapper.js", "srv\js\jquery_wrapper.js", 1)
FileInstall("srv\js\jspdf.plugin.autotable.min.js", "srv\js\jspdf.plugin.autotable.min.js", 1)
FileInstall("srv\js\jspdf.umd.min.js", "srv\js\jspdf.umd.min.js", 1)
FileInstall("srv\js\tabulator_esm.js", "srv\js\tabulator_esm.js", 1)
FileInstall("srv\js\tabulator_esm.js.map", "srv\js\tabulator_esm.js.map", 1)
FileInstall("srv\js\tabulator_esm.min.js", "srv\js\tabulator_esm.min.js", 1)
FileInstall("srv\js\tabulator_esm.min.js.map", "srv\js\tabulator_esm.min.js.map", 1)
FileInstall("srv\js\tabulator_esm.min.mjs", "srv\js\tabulator_esm.min.mjs", 1)
FileInstall("srv\js\tabulator_esm.min.mjs.map", "srv\js\tabulator_esm.min.mjs.map", 1)
FileInstall("srv\js\tabulator_esm.mjs", "srv\js\tabulator_esm.mjs", 1)
FileInstall("srv\js\tabulator_esm.mjs.map", "srv\js\tabulator_esm.mjs.map", 1)
FileInstall("srv\js\tabulator.js", "srv\js\tabulator.js", 1)
FileInstall("srv\js\tabulator.js.map", "srv\js\tabulator.js.map", 1)
FileInstall("srv\js\tabulator.min.js", "srv\js\tabulator.min.js", 1)
FileInstall("srv\js\tabulator.min.js.map", "srv\js\tabulator.min.js.map", 1)
FileInstall("srv\js\xlsx.full.min.js", "srv\js\xlsx.full.min.js", 1)

If ! DirExist("srv\css")
	DirCreate("srv\css")
FileInstall("srv\css\tabulator_bootstrap3.css", "srv\css\tabulator_bootstrap3.css", 1)
FileInstall("srv\css\tabulator_bootstrap3.css.map", "srv\css\tabulator_bootstrap3.css.map", 1)
FileInstall("srv\css\tabulator_bootstrap3.min.css", "srv\css\tabulator_bootstrap3.min.css", 1)
FileInstall("srv\css\tabulator_bootstrap3.min.css.map", "srv\css\tabulator_bootstrap3.min.css.map", 1)
FileInstall("srv\css\tabulator_bootstrap4.css", "srv\css\tabulator_bootstrap4.css", 1)
FileInstall("srv\css\tabulator_bootstrap4.css.map", "srv\css\tabulator_bootstrap4.css.map", 1)
FileInstall("srv\css\tabulator_bootstrap4.min.css", "srv\css\tabulator_bootstrap4.min.css", 1)
FileInstall("srv\css\tabulator_bootstrap4.min.css.map", "srv\css\tabulator_bootstrap4.min.css.map", 1)
FileInstall("srv\css\tabulator_bootstrap5.css", "srv\css\tabulator_bootstrap5.css", 1)
FileInstall("srv\css\tabulator_bootstrap5.css.map", "srv\css\tabulator_bootstrap5.css.map", 1)
FileInstall("srv\css\tabulator_bootstrap5.min.css", "srv\css\tabulator_bootstrap5.min.css", 1)
FileInstall("srv\css\tabulator_bootstrap5.min.css.map", "srv\css\tabulator_bootstrap5.min.css.map", 1)
FileInstall("srv\css\tabulator_bulma.css", "srv\css\tabulator_bulma.css", 1)
FileInstall("srv\css\tabulator_bulma.css.map", "srv\css\tabulator_bulma.css.map", 1)
FileInstall("srv\css\tabulator_bulma.min.css", "srv\css\tabulator_bulma.min.css", 1)
FileInstall("srv\css\tabulator_bulma.min.css.map", "srv\css\tabulator_bulma.min.css.map", 1)
FileInstall("srv\css\tabulator_materialize.css", "srv\css\tabulator_materialize.css", 1)
FileInstall("srv\css\tabulator_materialize.css.map", "srv\css\tabulator_materialize.css.map", 1)
FileInstall("srv\css\tabulator_materialize.min.css", "srv\css\tabulator_materialize.min.css", 1)
FileInstall("srv\css\tabulator_materialize.min.css.map", "srv\css\tabulator_materialize.min.css.map", 1)
FileInstall("srv\css\tabulator_midnight.css", "srv\css\tabulator_midnight.css", 1)
FileInstall("srv\css\tabulator_midnight.css.map", "srv\css\tabulator_midnight.css.map", 1)
FileInstall("srv\css\tabulator_midnight.min.css", "srv\css\tabulator_midnight.min.css", 1)
FileInstall("srv\css\tabulator_midnight.min.css.map", "srv\css\tabulator_midnight.min.css.map", 1)
FileInstall("srv\css\tabulator_modern.css", "srv\css\tabulator_modern.css", 1)
FileInstall("srv\css\tabulator_modern.css.map", "srv\css\tabulator_modern.css.map", 1)
FileInstall("srv\css\tabulator_modern.min.css", "srv\css\tabulator_modern.min.css", 1)
FileInstall("srv\css\tabulator_modern.min.css.map", "srv\css\tabulator_modern.min.css.map", 1)
FileInstall("srv\css\tabulator_semanticui.css", "srv\css\tabulator_semanticui.css", 1)
FileInstall("srv\css\tabulator_semanticui.css.map", "srv\css\tabulator_semanticui.css.map", 1)
FileInstall("srv\css\tabulator_semanticui.min.css", "srv\css\tabulator_semanticui.min.css", 1)
FileInstall("srv\css\tabulator_semanticui.min.css.map", "srv\css\tabulator_semanticui.min.css.map", 1)
FileInstall("srv\css\tabulator_simple.css", "srv\css\tabulator_simple.css", 1)
FileInstall("srv\css\tabulator_simple.css.map", "srv\css\tabulator_simple.css.map", 1)
FileInstall("srv\css\tabulator_simple.min.css", "srv\css\tabulator_simple.min.css", 1)
FileInstall("srv\css\tabulator_simple.min.css.map", "srv\css\tabulator_simple.min.css.map", 1)
FileInstall("srv\css\tabulator_site_dark.css", "srv\css\tabulator_site_dark.css", 1)
FileInstall("srv\css\tabulator_site_dark.css.map", "srv\css\tabulator_site_dark.css.map", 1)
FileInstall("srv\css\tabulator_site_dark.min.css", "srv\css\tabulator_site_dark.min.css", 1)
FileInstall("srv\css\tabulator_site_dark.min.css.map", "srv\css\tabulator_site_dark.min.css.map", 1)
FileInstall("srv\css\tabulator_site.css", "srv\css\tabulator_site.css", 1)
FileInstall("srv\css\tabulator_site.css.map", "srv\css\tabulator_site.css.map", 1)
FileInstall("srv\css\tabulator_site.min.css", "srv\css\tabulator_site.min.css", 1)
FileInstall("srv\css\tabulator_site.min.css.map", "srv\css\tabulator_site.min.css.map", 1)
FileInstall("srv\css\tabulator.css", "srv\css\tabulator.css", 1)
FileInstall("srv\css\tabulator.css.map", "srv\css\tabulator.css.map", 1)
FileInstall("srv\css\tabulator.min.css", "srv\css\tabulator.min.css", 1)
FileInstall("srv\css\tabulator.min.css.map", "srv\css\tabulator.min.css.map", 1)

TraySetIcon "main.ico"

Global CaddyResponse := ""

If FileExist("caddy.pid") {
	Loop Read, "caddy.pid" {
		Loop Parse, A_LoopReadLine, "`n" {
			Global CaddyPID := A_LoopField
		}
	Break
	}
	Else {
		Global CaddyPID := ""
	}
}

Global SavedTheme := IniRead("config.ini", "Theme", "Saved", "tabulator.css")
Global SkipLogin := IniRead("config.ini", "Login", "Skip", 1)

WHR := ComObject("WinHttp.WinHttpRequest.5.1")
WHR.Open("GET", "http://127.0.0.1:8080", true)
WHR.Send()
Try {
	WHR.WaitForResponse()
	Global CaddyResponse := WHR.ResponseText
}

If CaddyResponse == "" and RegExMatch(CaddyResponse, "tabulator") = 0 {
	Run A_ComSpec " /c `"" A_WorkingDir "\bin\win\amd64\caddy.exe`" file-server --listen :8080 --root " A_WorkingDir "\srv",, "Hide", &CaddyPID
	CaddyPIDFile := FileOpen("caddy.pid", "w")
	CaddyPIDFile.Write(CaddyPID)
	CaddyPIDFile.Close
}

ConnectionString := "Provider=MSOLEDBSQL;Data Source=(localdb)\MSSQLLocalDB;AttachDbFileName=" A_WorkingDir "\ReportsDemo.mdf;Initial Catalog=ReportsDemo;Integrated Security=SSPI;"

UserLoja := ""
UserPriv := ""
UserLogin := ""

ReportBrowser() {
	cssList := Array()
	ReportsList := Array()
	PrettyReportList := Array()
	ValidReportDir := ["relatórios", "relatorios", "reports"]
	ValidExtensions := [".txt", ".sql"]
	ReLoadFiles(*) {
		cssList := Array()
		ReportsList := Array()
		PrettyReportList := Array()
		Try {
			ddlReport.Delete
			ddlTheme.Delete
		}
		Loop Files A_WorkingDir "\srv\css\*.css" {
			cssList.Push(A_LoopFileName)
		}
		For K, V in ValidReportDir {
			For KExt, VExt in ValidExtensions {
				If DirExist(V) {
					Loop Files A_WorkingDir "\" V "\*" VExt {
						reportTitle := ""
						reportDescription := ""
						reportAuthor := ""
						SplitPath(A_LoopFileFullPath, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt, &OutDrive)
						If FileExist(OutDir "\" OutNameNoExt ".meta") {
							Loop Read, OutDir "\" OutNameNoExt ".meta" {
								Loop Parse, A_LoopReadLine, "`n" {
									If RegExMatch(A_LoopField, "title: ?") > 0 {
										reportTitle := RegExReplace(A_LoopField, "title: ?" , "")
									}
									If RegExMatch(A_LoopField, "description: ?") > 0 {
										reportDescription := RegExReplace(A_LoopField, "description: ?" , "")
									}
									If RegExMatch(A_LoopField, "author: ?") > 0 {
										reportAuthor := RegExReplace(A_LoopField, "author: ?" , "")
									}
								}	
							}
						}
						ReportsList.Push(A_LoopFileFullPath "|" OutDir "|" OutFileName "|" OutExtension "|" OutNameNoExt "|" reportTitle "|" reportDescription "|" reportAuthor)
					}
				}
			}
		}
		For K1, V1 in ReportsList {
			For K2, V2 in StrSplit(V1, "|") {
				If K2 = 6 and V2 != "" {
					PrettyReportList.Push(V2)
				}
				Else If K2 = 5 and StrSplit(V1, "|")[6] == "" {
					PrettyReportList.Push(V2)
				}
			}
		}
		Try {
			ddlReport.Add(PrettyReportList)
			ddlTheme.Add(cssList)
			ddlTheme.Choose(SavedTheme)
			FavThemeLoader()
		}
	}
	FavThemeLoader(*) {
		If ddlTheme.Text == SavedTheme {
			btnFavTheme.Text := "★"
		} Else {
			btnFavTheme.Text := "☆"
		}
	}
	FavThemeChanger(*) {
		Global SavedTheme := ddlTheme.Text
		IniWrite(SavedTheme, "demo_reports.ini", "Theme", "Saved")
		FavThemeLoader()
	}
	ReLoadFiles()
	Main := Gui("", "Reports - Browser")
	grpBoxMain := Main.Add("GroupBox", "Hidden W300 H100", "")
	grpBoxMain.GetPos(&grpBoxMainX, &grpBoxMainY, &grpBoxMainW, &grpBoxMainH)
	txtReport := Main.Add("Text", "X" grpBoxMainX+10 " Y" grpBoxMainY, "Relatório: ")
	txtReport.GetPos(&txtReportX, &txtReportY, &txtReportW, &txtReportH)
	ddlReport := Main.Add("DropDownList", "W259", PrettyReportList)
	ddlReport.GetPos(&ddlReportX, &ddlReportY, &ddlReportW, &ddlReportH)
	btnReportInfo := Main.Add("Button", "X" ddlReportW+ddlReportX+2 " Y" ddlReportY-1 " W20 H" ddlReportH+2, "ℹ")
	btnReportInfo.OnEvent("Click", ReportInfo)
	chkDataInicial := Main.Add("Checkbox", "X" ddlReportX " Y" ddlReportY+ddlReportH+5, "Data Inicial:")
	chkDataInicial.OnEvent("Click", dtToggle)
	chkDataInicial.GetPos(&chkDataInicialX, &chkDataInicialY, &chkDataInicialW, &chkDataInicialH)
	dtDataInicial := Main.Add("DateTime", "W90", "ShortDate")
	dtDataInicial.GetPos(&dtDataInicialX, &dtDataInicialY, &dtDataInicialW, &dtDataInicialH)
	dtDataInicial.Value := ""
	dtDataInicial.Opt("+Disabled")
	chkDataFinal := Main.Add("Checkbox", "X" dtDataInicialX+dtDataInicialW+7 " Y" chkDataInicialY, "Data Final:")
	chkDataFinal.OnEvent("Click", dtToggle)
	dtDataFinal := Main.Add("DateTime", "W90", "ShortDate")
	dtDataFinal.GetPos(&dtDataFinalX, &dtDataFinalY, &dtDataFinalW, &dtDataFinalH)
	dtDataFinal.Value := ""
	dtDataFinal.Opt("+Disabled")
	btnGenerate := Main.Add("Button", "X" dtDataFinalW+dtDataFinalX+2 " Y" dtDataFinalY-1, "→")
	btnGenerate.GetPos(&btnGenerateX, &btnGenerateY, &btnGenerateW, &btnGenerateH)
	btnGenerate.OnEvent("Click", ReportGenerator)
	btnReload := Main.Add("Button", "X" btnGenerateW+btnGenerateX+2 " Y" btnGenerateY, "⟳")
	btnReload.OnEvent("Click", ReLoadFiles)
	txtTheme := Main.Add("Text", "X" grpBoxMainX+10 " Y" dtDataFinalY+dtDataFinalH+5, "Tema: ")
	ddlTheme := Main.Add("DropDownList", "X" ddlReportX " W259", cssList)
	ddlTheme.GetPos(&ddlThemeX, &ddlThemeY, &ddlThemeW, &ddlThemeH)
	ddlTheme.OnEvent("Change", FavThemeLoader)
	ddlTheme.Choose(SavedTheme)
	btnFavTheme := Main.Add("Button", "X" ddlThemeW+ddlThemeX+2 " Y" ddlThemeY-1 " W20 H" ddlThemeH+2, "☆")
	btnFavTheme.OnEvent("Click", FavThemeChanger)
	sbMain := Main.Add("StatusBar", "", "")
	Main.Show("Hide")
	Main.Show("W319")
	FavThemeLoader()
	dtToggle(*) {
		If chkDataInicial.Value = 1 {
			dtDataInicial.Opt("-Disabled")
		} Else {
			dtDataInicial.Opt("+Disabled")
		}
		If chkDataFinal.Value = 1 {
			dtDataFinal.Opt("-Disabled")
		} Else {
			dtDataFinal.Opt("+Disabled")
		}
	}
	ReportInfo(*) {
		If ddlReport.Value > 0 {
			Msgbox "Título: " StrSplit(ReportsList[ddlReport.Value], "|")[6] "`nDescrição: " StrSplit(ReportsList[ddlReport.Value], "|")[7] "`nAutor: " StrSplit(ReportsList[ddlReport.Value], "|")[8], "Informações Adicionais", "Iconi"
		}
	}
	ReportGenerator(*) {
		If ddlReport.Value = 0 {
			MsgBox "Nenhum relatório selecionado!", "Aviso", "Icon!"
			Return
		}
		sbMain.SetText("Aguarde...")
		If !(chkDataInicial.Value)
			DataInicialRelatorio := "00:00:00"
		Else
			DataInicialRelatorio := FormatTime(dtDataInicial.Value, "yyyy-MM-dd 00:00:00")
		If !(chkDataFinal.Value)
			DataFinalRelatorio := "00:00:00"
		Else
			DataFinalRelatorio := FormatTime(dtDataFinal.Value, "yyyy-MM-dd 00:00:00")
		ReportQueryFile := StrSplit(ReportsList[ddlReport.Value], "|")[1]
		ReportQuery := FileRead(ReportQueryFile)
		SplitPath(ReportQueryFile, &ReportQueryFileName, &ReportQueryFileDir, &ReportQueryFileExt, &ReportQueryFileName, &ReportQueryFileDrive)
		ReportQuery := StrReplace(ReportQuery, "`{`{DATAINICIAL`}`}", DataInicialRelatorio)
		ReportQuery := StrReplace(ReportQuery, "`{`{DATAFINAL`}`}", DataFinalRelatorio)
		If (UserLoja)
			ReportQuery := StrReplace(ReportQuery, "`{`{CODLOJA`}`}", UserLoja)
		If (UserPriv)
			ReportQuery := StrReplace(ReportQuery, "`{`{PRIV`}`}", UserPriv => "0")
		If (UserLogin)
			ReportQuery := StrReplace(ReportQuery, "`{`{LOGIN`}`}", UserLogin => "0")
		adoQuery := ADOSQL(ConnectionString, ReportQuery)
		jsonData := "[`n"
		Loop adoQuery.Length {
			If A_Index > 1 {			
				jsonData .= "{"
				For K, V in adoQuery[A_Index] {
					If K > 1
						jsonData .= ", "
					jsonData .= "`"" adoQuery[1][K] "`"" ": " "`"" V "`""
				}
				if A_Index = adoQuery.Length
					jsonData .= "}`n"
				else
					jsonData .= "},`n"
			}
		}
		jsonData .= "]"
		html_open := "
		(
		<html>
		)"
		html_head_open := "
		(
			<head>
		)"
		html_head := "
		(
				<title>demo Report</title>
				<link rel="icon" type="image/x-icon" href="../images/favicon.ico"> </link>
				<link href="../css/{1}" rel="stylesheet"> </link>
				<link rel="stylesheet" href="../fontawesome-free/5.15.4/css/all.css" </link>
				<script type="text/javascript" src="../js/xlsx.full.min.js"></script>
				<script src="../js/jspdf.umd.min.js"></script>
				<script src="../js/jspdf.plugin.autotable.min.js"></script>
				<script src="../js/tabulator.js"> </script>
		)"
		html_head_close := "
		(
			</head>
		)"
		html_body_open := "
		(
			<body>
		)"
		html_body := "
		(
				<div>
					<button id="download-csv">CSV</button>
					<button id="download-json">JSON</button>
					<button id="download-xlsx">XLSX</button>
					<button id="download-pdf">PDF</button>
					<button id="download-html">HTML</button>
				</div>
				<div id="report"></div>
		)"
		html_style_open := "
		(
			<style>
		)"
		html_style := "
		(
		)"
		html_style_close := "
		(
			</style>
		)"
		html_script_open := "
		(
				<script>
		)"
		html_script := "
		(
					var table = new Tabulator("#report", {
						ajaxURL: "data.json",
						autoColumns:"full",
						autoResize:false,
						layout:"fitData",
						movableColumns:true,
						columnDefaults:{
							tooltip:true,
							headerFilter:"list",
								headerFilterParams: {
									valuesLookup:true,
									clearable:true,
									multiselect: true
								},
						}
					});
					document.getElementById("download-csv").addEventListener("click", function(){
						table.download("csv", "data.csv");
					});
					document.getElementById("download-json").addEventListener("click", function(){
						table.download("json", "data.json");
					});
					document.getElementById("download-xlsx").addEventListener("click", function(){
						table.download("xlsx", "data.xlsx", {sheetName:"Report"});
					});
					document.getElementById("download-pdf").addEventListener("click", function(){
						table.download("pdf", "data.pdf", {
							orientation:"landscape",
							title:"Report",
						});
					});
					document.getElementById("download-html").addEventListener("click", function(){
						table.download("html", "data.html", {style:true});
					});
		)"
		html_script_close := "
		(
				</script>
		)"
		html_body_close := "
		(
			</body>
		)"
		html_close := "
		(
		</html>
		)"
		If FileExist("__default.html")
			html_body := FileRead("__default.html")
		If FileExist("__default.js")
			html_body := FileRead("__default.js")
		If FileExist("__default.css")
			html_style := FileRead("__default.css")
		If FileExist(ReportQueryFileDir "\" ReportQueryFileName ".html")
			html_body := FileRead(ReportQueryFileDir "\" ReportQueryFileName ".html")
		If FileExist(ReportQueryFileDir "\" ReportQueryFileName ".js" )
			html_script := FileRead(ReportQueryFileDir "\" ReportQueryFileName ".js")
		If FileExist(ReportQueryFileDir "\" ReportQueryFileName ".css" )
			html_style := FileRead(ReportQueryFileDir "\" ReportQueryFileName ".css")
		html := html_open "`n" html_head_open "`n" Format(html_head, ddlTheme.Text) "`n" html_head_close "`n" html_body_open "`n" html_body "`n" html_style_open "`n" html_style "`n" html_style_close "`n" html_script_open "`n" html_script "`n" html_script_close "`n" html_body_close "`n" html_close
		If FileExist("srv\data.json")
			FileDelete("srv\data.json")
		If FileExist("srv\index.html")
			FileDelete("srv\index.html")
		jsonDataFile := FileOpen("srv\data.json", "w", "UTF-8")
		jsonDataFile.Write(jsonData)
		jsonDataFile.Close
		tabulatorFile := FileOpen("srv\index.html", "w", "UTF-8")
		tabulatorFile.Write(html)
		tabulatorFile.Close
		sbMain.SetText("")
		Run("http://127.0.0.1:8080")
	}
}

ReportBrowser
