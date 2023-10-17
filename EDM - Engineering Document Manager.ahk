#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance,Force

;...................................................................Zmienne podstawoe..............................................................

FileEncoding, UTF-8
info_nazwa_programu := "EDM - Engineering Document Manager"
info_wersja_programu := "0.98"

info_DMS = %A_ScriptDir%\Info_DMS.txt
If !FileExist(info_DMS)
{
	Dane2 = %A_ScriptDir%\Dane
	Roboczy2 = %A_ScriptDir%\Roboczy
	Wzor2 = %A_ScriptDir%\Wzor
	Repozytorium2 = %A_ScriptDir%\Rep
	Dokumentacja2 = %A_ScriptDir%\Dokumentacja
	ukryc_DMS = 1
	SplitPath, A_ScriptDir,,dir_projekty,,Nazwa_projektu2
	Gosub, dodaj_dr_pr
	WinWaitClose, Tworzenie nowego projektu
}

FileReadLine, Dane, %A_ScriptDir%\Info_DMS.txt, 1
FileReadLine, Roboczy, %A_ScriptDir%\Info_DMS.txt, 2
FileReadLine, Wzor, %A_ScriptDir%\Info_DMS.txt, 3
FileReadLine, Repozytorium, %A_ScriptDir%\Info_DMS.txt, 4
FileReadLine, Dokumentacja, %A_ScriptDir%\Info_DMS.txt, 5
FileReadLine, dir_projekty, %A_ScriptDir%\Info_DMS.txt, 6

Dane3 = %Dane%\Branza
Dane4 = %Dane%\Licznik
Dane5 = %Dane%\Rodzaje_dok
Dane6 = %Dane%\Uprawnienia
Dane7 = %Dane%\Temp

If !FileExist(Dane)
	FileCreateDir, %Dane%
If !FileExist(Dane3)
	FileCreateDir, %Dane3%
If !FileExist(Dane4)
	FileCreateDir, %Dane4%
If !FileExist(Dane5)
	FileCreateDir, %Dane5%
If !FileExist(Dane6)
	FileCreateDir, %Dane6%
If !FileExist(Dane7)
	FileCreateDir, %Dane7%
If !FileExist(Dokumentacja)
	FileCreateDir, %Dokumentacja%
If !FileExist(Roboczy)
	FileCreateDir, %Roboczy%
If !FileExist(Wzor)
	FileCreateDir, %Wzor%
If !FileExist(Repozytorium)
	FileCreateDir, %Repozytorium%

Licznik_dir = %Dane%\Licznik
Licznik2_dir = %Licznik_dir%\info.txt
IF !FileExist(Licznik_dir)
{
	FileCreateDir, %Licznik_dir%
	FileAppend, 00`n001`n-, %Licznik_dir%\info.txt
}
If !FileExist(Licznik2_dir)
	FileAppend, 00`n001`n-, %Licznik_dir%\info.txt
FileReadLine, Licznik1, %Licznik_dir%\info.txt, 1
FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2
FileReadLine, Licznik3, %Licznik_dir%\info.txt, 3

upr1_uzytkownika = %Dane%\Uprawnienia\%A_UserName%.txt
If !FileExist(upr1_uzytkownika)
	FileAppend, Pełne`n%A_UserName%, %upr1_uzytkownika%
FileReadLine, upr_uzytkownika, %upr1_uzytkownika%, 1
FileReadLine, upr_imie, %upr1_uzytkownika%, 2
If (upr_uzytkownika = "")
	upr_uzytkownika := "Pełne"

;...................................................................Glowne menu programu..............................................................

Gui, dok:+HwndGuiHwnd
Gui, dok:Font, Bold
Gui, dok:Add, Text, x10 y25 w180 h20 Section, Numer dokumentacji:
Gui, dok:Font,
Gui, dok:Add, ListBox, y+5 w320 h140 vNumer_dok_2 gNumer_dok_2, %Numer_dok_2%
Gui, dok:Add, Button, y+5 w60 h20 vButtonToHide3 gNowy_pro, Nowy
Gui, dok:Add, Button, x+5 w60 h20 vButtonToHide4 gEdytuj_pro, Edytuj
Gui, dok:Add, Button, x+5 w60 h20 vButtonToHide1 gUsun_pro, Usuń
Gui, dok:Add, Button, x+5 w60 h20 vMoveButton1 gWydaj_pro, Wydaj
Gui, dok:Add, Button, x+5 w60 h20 vMoveButton2 gRaport_pro, Raport
Gui, dok:Add, Button, y+5 xs w60 h20 gOdswiez, Odśwież
Gui, dok:Add, Button, x+70 w60 h20 gPrzelacz, Przełącz
Gui, dok:Add, Button, x+5 w60 h20 gUprawnienia, Ustal
Gui, dok:Add, Button, x+5 w60 h20 ginfo_program, Info
Gui, dok:Font, Bold
Gui, dok:Add, Text, y265 xs h20 w120, Dokumenty:
Gui, dok:Font,
Gui, dok:Add, Text, x+10 Right w195 h20 vfiltr_stat, %filtr_stat%
Gui, dok:Add, ListBox, y+5 xs w320 h325 vdok_2 ginfo_dok_2, %dok_2%
Gui, dok:Add, Button, y+5 w60 h20 gotworz_dok_2, Otwórz
Gui, dok:Add, Button, x+5 w60 h20 godczyt_dok_2, Odczyt
Gui, dok:Add, Button, x+5 w60 h20 gNowy_dok, Nowy
Gui, dok:Add, Button, x+5 w60 h20 gimport_dok_2, Import
Gui, dok:Add, Button, x+5 w60 h20 gedytuj_dok_2, Edytuj
Gui, dok:Add, Button, y+5 xs w60 h20 gwymien_dok_2, Wymień
Gui, dok:Add, Button, x+5 w60 h20 gusun_dok_2, Usuń
Gui, dok:Add, Button, x+5 w60 h20 gArchiwizuj_dok, Archiwizuj
Gui, dok:Add, Button, x+5 w60 h20 gFiltr, Filtr

Gui, dok:Add, GroupBox, y5 x340 w670 h230 Section, Informacje o projekcie i dokumentacji:
Gui, dok:Font, Bold
Gui, dok:Add, Text, ys+20 xs+10 w250 h20, Informacja o projekcie:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 vButtonToHide2 ged_pro, Edytuj
Gui, dok:Add, Edit, y+5 xs+10 w320 h175 ReadOnly vInfo_projekt, %Info_projekt%
Gui, dok:Font, Bold
Gui, dok:Add, Text, ys+20 xs+340 h20 w250, Informacja o dokumentacji:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 vButtonToHide5 gedytuj_dok, Edytuj
Gui, dok:Add, Edit, y+5 xs+340 w320 h115 ReadOnly vInfo_dok, %Info_dok%
Gui, dok:Font, Bold
Gui, dok:Add, Text, y180 xs+340 h20 w250, Status dokumentacji:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 vButtonToHide6 gzmien2_dok, Zmień
Gui, dok:Add, Edit, y+5 xs+340 w320 h20 ReadOnly vStatus_pro, %Status_pro%

Gui, dok:Add, GroupBox, y245 x340 Section w670 h425, Informacje o dokumencie:
Gui, dok:Font, Bold
Gui, dok:Add, Text, y265 xs+10 h20, Pliki PDF:
Gui, dok:Font,
Gui, dok:Add, ListBox, y+5 w320 h100 vdok_pdf_2 gpdf_otworz_dok, Wybierz plik PDF||%dok_pdf_2%
Gui, dok:Add, Button, y+5 w60 h20 gotworz_pdf, Otwórz
Gui, dok:Add, Button, x+5 w60 h20 gdodaj_pdf, Dodaj
Gui, dok:Add, Button, x+5 w60 h20 gusun_pdf, Usuń
Gui, dok:Add, Button, x+5 w60 h20 gwymien_pdf, Wymień
Gui, dok:Font, Bold
Gui, dok:Add, Text, y425 xs+10 h20 w250, ETransmit:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 gzmien_etr, Zmień
Gui, dok:Add, Edit, y+5 xs+10 w320 h20 ReadOnly vetransmit_2, %etransmit_2%
Gui, dok:Font, Bold
Gui, dok:Add, Text, y485 xs+10 h20 w185, Archiwum dokumentu:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 gOtworz_arch, Otwórz
Gui, dok:Add, Button, x+5 w60 h20 gPrzywroc_arch, Przywróć
Gui, dok:Add, ListBox, y+5 xs+10 w320 h155 garch_otworz_dok varch_dok_2, %arch_dok_2%

Gui, dok:Font, Bold
Gui, dok:Add, Text, xs+340 y265 h20 w250, Opis dokumentu:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 gedytuj_opis_dok, Edytuj
Gui, dok:Add, Edit, y+5 xs+340 w320 h20 ReadOnly vOpis_dok, %Opis_dok%	
Gui, dok:Font, Bold
Gui, dok:Add, Text, xs+340 y+10 h20 w250, Rewizja:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 gzmien_rev, Zmień
Gui, dok:Add, Edit, y+5 xs+340 w320 h20 ReadOnly vRewizja_dok, %Rewizja_dok%
Gui, dok:Font, Bold
Gui, dok:Add, Text, y+10 xs+340 h20 w250, Status:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 gzmien_dok, Zmień
Gui, dok:Add, Edit, y+5 xs+340 w320 h20 ReadOnly vStatus_dok, %Status_dok%
Gui, dok:Font, Bold
Gui, dok:Add, Text, y+10 xs+340 h20 w185, Historia:
Gui, dok:Font,
Gui, dok:Add, Button, x+10 w60 h20 gNotka_dok, Notka
Gui, dok:Add, Button, x+5 w60 h20 gzalacznik, Załącznik
Gui, dok:Add, Edit, y+5 xs+340 w320 h200 ReadOnly vHis_dok, %His_dok%	

If else (upr_uzytkownika = "Rozszerzone")
{
	GuiControl,dok:Hide,ButtonToHide1
	GuiControl,dok:Hide,ButtonToHide2
	GuiControl,dok:Move, MoveButton1, x140
	GuiControl,dok:Move, MoveButton2, x205
}
If else (upr_uzytkownika = "Ograniczone")
{
	GuiControl,dok:Hide,ButtonToHide1
	GuiControl,dok:Hide,ButtonToHide2
	GuiControl,dok:Hide,ButtonToHide3
	GuiControl,dok:Hide,ButtonToHide4
	GuiControl,dok:Hide,ButtonToHide5
	GuiControl,dok:Hide,ButtonToHide6
	GuiControl,dok:Move, MoveButton1, x10
	GuiControl,dok:Move, MoveButton2, x75
	
}


FileRead, Info_projekt, %Repozytorium%/info.txt
GuiControl, dok: , Info_projekt, %Info_projekt% 

Numer_dok_2 := "Wybierz dokumentację||"
GuiControl, dok: , Numer_dok_2, |%Numer_dok_2% 
Loop, files, %Repozytorium%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	GuiControl, dok: , Numer_dok_2, %Numer_dok_2% 
	Folder_wyjscowy = %Roboczy%\%Numer_dok%
	If !FileExist(Folder_wyjscowy)
		FileCreateDir, %Folder_wyjscowy%
} 
Numer_dok_2 := "Wybierz dokumentację"

dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2% 
dok_2 := "Wybierz dokument"

SplitPath, A_ScriptDir,,,,Tytul_projektu

Gui, dok:Show, w1020 h680,%Tytul_projektu%
SetTimer, labelToolTips, 500

#If WinActive("ahk_id" GuiHwnd)
;Enter::
If (dok_2 != "Wybierz dokument")
{
	If (dok_pdf_2 != "Wybierz plik PDF")
		Gosub, otworz_pdf
	else if (arch_dok_2 != "")
		Gosub, Otworz_arch
	else
		Gosub, otworz_dok_2	
}
#If
return

#If WinActive("ahk_id" GuiHwnd)
del::
If (dok_2 != "Wybierz dokument")
{
	If (dok_pdf_2 != "Wybierz plik PDF")
		Gosub, usun_pdf
	else
		Gosub, usun_dok_2			
}
#If
return


;...................................................................funkcje - Edycja..............................................................


;_______________________________________________________Numer_dokumentacj_____________________________________________

Numer_dok_2:
Gui, dok:Submit, noHide

FileRead, Info_dok, %Repozytorium%\%Numer_dok_2%\Info.txt
GuiControl, dok: , Info_dok, %Info_dok%

dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%

inf_2_spr := ""

If (Numer_dok_2 = "Wybierz dokumentację")
{
	Status_pro := "" 
	GuiControl, dok:, Status_pro, %Status_pro%
	Rewizja_dok := ""
	GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
	Status_dok := ""
	GuiControl, dok: , Status_dok, %Status_dok% 	
	His_dok := ""
	GuiControl, dok: , His_dok, %His_dok% 	
	dok_pdf_2 := "Wybierz plik PDF||"
	GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
	etransmit_2 := ""
	GuiControl, dok: , etransmit_2, %etransmit_2% 
	arch_dok_2 := ""
	GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
	Opis_dok := ""
	GuiControl, dok: , Opis_dok, %Opis_dok% 
	return
}
else
{
	Status_pro := ""
	Status1_pro := ""
	Status2_pro := ""
	Status3_pro := ""
	done = 0
	Loop, files, %Repozytorium%\%Numer_dok_2%\*, D
	{ 
		FileReadLine, Spr_stat2, %A_LoopFileFullPath%\info.txt, 6
		If (Spr_stat2 = "Wydany")
			Status1_pro := "Wydana"
		If else (Spr_stat2 = "Akceptacja")
			Status2_pro := "Akceptacja"
		If else (Spr_stat2 != "Wydany" and Spr_stat2 != "Akceptacja")
			Status3_pro := "Roboczy"
	} 
	If (Status1_pro != "")
		Status_pro = %Status1_pro%
	If else (Status2_pro != "")
		Status_pro = %Status2_pro%
	If else (Status3_pro != "")
		Status_pro = %Status3_pro%
	
	GuiControl, dok: , Status_pro, %Status_pro%	
}

dok_2 := "Wybierz dokument"
Loop, files, %Repozytorium%\%Numer_dok_2%\*%filtr2%*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,dok_2
	FileReadLine, dok_2_spr, %A_LoopFileFullPath%\info.txt, 1
	FileReadLine, Spr_stat, %A_LoopFileFullPath%\info.txt, 6
	If (filtr3 = "")
		GuiControl, dok: , dok_2, %dok_2%
	If else (Filtr3 = Spr_stat)
		GuiControl, dok: , dok_2, %dok_2%
	
	dok_2_spr = %Roboczy%\%Numer_dok_2%\%dok_2_spr%
	If !FileExist(dok_2_spr)
		inf_2_spr = %inf_2_spr%%dok_2% `n	
} 
dok_2 := "Wybierz dokument"

If (inf_2_spr != "")
	MsgBox, 16, Ostrzeżenie, Niezgodność w pliku:`n%inf_2_spr% 

Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

return

Uprawnienia:
Gui, uprawnienia:Destroy
Gui, dok:Submit, noHide

upr1_uzytkownika = %Dane%\Uprawnienia\%A_UserName%.txt
FileReadLine, upr_uzytkownika, %Dane%\Uprawnienia\%A_UserName%.txt, 1

If (upr_uzytkownika = "Pełne")
{
	upr1 = 1
	upr2 = 0
	upr3 = 0
}
If (upr_uzytkownika = "Rozszerzone")
{
	upr1 = 0
	upr2 = 1
	upr3 = 0
}
If (upr_uzytkownika = "Ograniczone")
{
	upr1 = 0
	upr2 = 0
	upr3 = 1
}

Gui, uprawnienia:Font, Bold
Gui, uprawnienia:Add, Text, x10 y10 h20 w270 Section, Aktualne uprawnienia użytkownika: %A_UserName%
Gui, uprawnienia:Font,
Gui, uprawnienia:Add, Radio, y+5 xs=10 vuprawnienia_radio checked%upr1% Group, Pełne
Gui, uprawnienia:Add, Radio, checked%upr2%, Rozszerzone
Gui, uprawnienia:Add, Radio, checked%upr3%, Ograniczone
Gui, uprawnienia:Font, Bold
Gui, uprawnienia:Add, Text, y+10 xs h20, Imię i Nazwisko użytkownika:
Gui, uprawnienia:Font,
Gui, uprawnienia:Add, Edit, y+5 w270 vupr_imie, %upr_imie%
Gui, uprawnienia:Add, Button, y+10 xs Default w60 gzmien_uprawnienia, Zmień
Gui, uprawnienia:Add, Button, x+10 w60 vButtonToHide1 guprawnienia_grupy, Grupy	

If else (upr_uzytkownika = "Rozszerzone" or upr_uzytkownika = "Ograniczone")
{
	GuiControl,uprawnienia:Hide,ButtonToHide1
}

Gui, uprawnienia:Show, , Uprawnienia
return

zmien_uprawnienia:
Gui, uprawnienia:Submit,Nohide

If (uprawnienia_radio=1)
	New_upr := "Pełne"
If (uprawnienia_radio=2)
	New_upr := "Rozszerzone"
If (uprawnienia_radio=3)
	New_upr := "Ograniczone"
IF (upr_imie = "")
	upr_imie = %A_UserName%

FileDelete, %upr1_uzytkownika%
FileAppend, %New_upr%`n%upr_imie%, %upr1_uzytkownika%

Gui, uprawnienia:Destroy
Reload
return

uprawnienia_grupy:
Gui, uprawnienia2:Destroy
Gui, uprawnienia:Submit,Nohide

Gui, uprawnienia2:Font, Bold
Gui, uprawnienia2:Add, Text, x10 y10 h20 w200 Section, Dostępni użytkownicy::
Gui, uprawnienia2:Font,
Gui, uprawnienia2:Add, ListBox, y+5 w200 h365 vgrupa_upr_0, %grupa_upr_0%
Gui, uprawnienia2:Font, Bold
Gui, uprawnienia2:Add, Text, x220 y10 h20 w200 Section, Uprawnienia Pełne:
Gui, uprawnienia2:Font,
Gui, uprawnienia2:Add, ListBox, y+5 w200 h100 ReadOnly vgrupa_upr_1, %grupa_upr_1%
Gui, uprawnienia2:Font, Bold
Gui, uprawnienia2:Add, Text, y+10 h20 w200, Uprawnienia Rozszerzone:
Gui, uprawnienia2:Font,
Gui, uprawnienia2:Add, ListBox, y+5 w200 h100 ReadOnly vgrupa_upr_2, %grupa_upr_2%
Gui, uprawnienia2:Font, Bold
Gui, uprawnienia2:Add, Text, y+10 h20 w200 ReadOnly, Uprawnienia Ograniczone:
Gui, uprawnienia2:Font,
Gui, uprawnienia2:Add, ListBox, y+5 w200 h100 vgrupa_upr_3, %grupa_upr_3%

Gui, uprawnienia2:Add, Button, x10 y400 w60 Default Section guprawnienia2_grupy, Dodaj
Gui, uprawnienia2:Add, Button, x+10 w60 guprawnienia4_grupy, Edytuj
Gui, uprawnienia2:Add, Button, x+10 w60 guprawnienia5_grupy, Usuń

grupa_upr_0 := ""
GuiControl, uprawnienia2:, grupa_upr_0, |%grupa_upr_0%
grupa_upr_1 := ""
GuiControl, uprawnienia2:, grupa_upr_1, |%grupa_upr_1%
grupa_upr_2 := ""
GuiControl, uprawnienia2:, grupa_upr_2, |%grupa_upr_2%
grupa_upr_3 := ""
GuiControl, uprawnienia2:, grupa_upr_3, |%grupa_upr_3%

Loop, files, %Dane%\Uprawnienia\*, R
{ 
	SplitPath, A_LoopFileFullPath,,,,grupa_upr_0
	GuiControl, uprawnienia2:, grupa_upr_0, %grupa_upr_0%
	
	FileReadLine, grupa_upr_spr, %A_LoopFileFullPath%, 1
	If (grupa_upr_spr = "Pełne")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_1
		GuiControl, uprawnienia2:, grupa_upr_1, %grupa_upr_1%
	}
	If (grupa_upr_spr = "Rozszerzone")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_2
		GuiControl, uprawnienia2:, grupa_upr_2, %grupa_upr_2%
	}
	If (grupa_upr_spr = "Ograniczone")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_3
		GuiControl, uprawnienia2:, grupa_upr_3, %grupa_upr_3%
	}
} 
grupa_upr_1 := ""
grupa_upr_2 := ""
grupa_upr_3 := ""
grupa_upr_0 := ""

Gui, uprawnienia2:Show, , Grupy uprawnień
return

uprawnienia5_grupy:
Gui, uprawnienia2:Submit, noHide
Gui, uprawnienia5:Destroy

If (grupa_upr_0 = "")
{
	MsgBox, 16, Ostrzeżenie, Wybierz uprawnienia
	return
}

Gui, uprawnienia5:Add, Text, x10 y10 h20 w270, Czy na pewno chcesz usunąć zaznaczone uprawnienia?
Gui, uprawnienia5:Add, Button, y+10 x70 w60 guprawnienia6_grupy, Usuń
Gui, uprawnienia5:Add, Button, x+10 w60 Default guprawnienia7_grupy , Anuluj

Gui, uprawnienia5:Show,, Ostrzeżenie!
return

uprawnienia6_grupy:
upr_dir_uzytkownika = %Dane%\Uprawnienia\%grupa_upr_0%.txt
FileDelete, %upr_dir_uzytkownika%

grupa_upr_0 := ""
GuiControl, uprawnienia2:, grupa_upr_0, |%grupa_upr_0%
grupa_upr_1 := ""
GuiControl, uprawnienia2:, grupa_upr_1, |%grupa_upr_1%
grupa_upr_2 := ""
GuiControl, uprawnienia2:, grupa_upr_2, |%grupa_upr_2%
grupa_upr_3 := ""
GuiControl, uprawnienia2:, grupa_upr_3, |%grupa_upr_3%

Loop, files, %Dane%\Uprawnienia\*, R
{ 
	SplitPath, A_LoopFileFullPath,,,,grupa_upr_0
	GuiControl, uprawnienia2:, grupa_upr_0, %grupa_upr_0%
	
	FileReadLine, grupa_upr_spr, %A_LoopFileFullPath%, 1
	If (grupa_upr_spr = "Pełne")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_1
		GuiControl, uprawnienia2:, grupa_upr_1, %grupa_upr_1%
	}
	If (grupa_upr_spr = "Rozszerzone")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_2
		GuiControl, uprawnienia2:, grupa_upr_2, %grupa_upr_2%
	}
	If (grupa_upr_spr = "Ograniczone")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_3
		GuiControl, uprawnienia2:, grupa_upr_3, %grupa_upr_3%
	}
} 
grupa_upr_1 := ""
grupa_upr_2 := ""
grupa_upr_3 := ""
grupa_upr_0 := ""

Gui, uprawnienia5:Destroy
return

uprawnienia7_grupy:
Gui, uprawnienia5:Destroy
return

uprawnienia4_grupy:
Gui, uprawnienia2:Submit,Nohide
Gui, uprawnienia4:Destroy

upr_dir2_uzytkownika = %Dane%\Uprawnienia\%grupa_upr_0%.txt
FileReadLine, upr_dir_uzytkownika, %upr_dir2_uzytkownika%, 1
FileReadLine, grupa_upr_imie, %upr_dir2_uzytkownika%, 2
grupa_0_upr =%Dane%\Uprawnienia\%grupa_upr_0%.txt

If (upr_dir_uzytkownika = "Pełne")
	grupa_upr_4 = Pełne||Rozszerzone|Ograniczone
If (upr_dir_uzytkownika = "Rozszerzone")
	grupa_upr_4 = Pełne|Rozszerzone||Ograniczone
If (upr_dir_uzytkownika = "Ograniczone")
	grupa_upr_4 = Pełne|Rozszerzone|Ograniczone||

Gui, uprawnienia4:Add, Text, x10 y10 h20 w150 Section, Nazwa użytkownika Windows:
Gui, uprawnienia4:Add, Edit, x+10 w110 vgrupa_upr_nazwa, %grupa_upr_0%
Gui, uprawnienia4:Add, Text, y+10 xs h20 w150 , Grupa uprawnień:
Gui, uprawnienia4:Add, DropDownList, x+10 w110 vgrupa_upr_4, %grupa_upr_4%
Gui, uprawnienia4:Add, Text, y+10 xs h20 w150 , Imię i Nazwisko użytkownika:
Gui, uprawnienia4:Add, Edit, x+10 w150 vgrupa_upr_imie, %grupa_upr_imie%
Gui, uprawnienia4:Add, Button, y+10 xs Default w60 guprawnienia3_grupy, Zmień

Gui, uprawnienia4:Show, , Dodaj użytkownika
return

uprawnienia2_grupy:
Gui, uprawnienia3:Destroy

grupa_upr_imie := ""
grupa_upr_nazwa := ""
Gui, uprawnienia3:Add, Text, x10 y10 h20 w150 Section, Nazwa użytkownika Windows:
Gui, uprawnienia3:Add, Edit, x+10 w150 vgrupa_upr_nazwa, 
Gui, uprawnienia3:Add, Text, y+10 xs h20 w150 , Grupa uprawnień:
Gui, uprawnienia3:Add, DropDownList, x+10 w150 vgrupa_upr_4, Pełne||Rozszerzone|Ograniczone
Gui, uprawnienia3:Add, Text, y+10 xs h20 w150 , Imię i Nazwisko użytkownika:
Gui, uprawnienia3:Add, Edit, x+10 w150 vgrupa_upr_imie, 
Gui, uprawnienia3:Add, Button, y+10 xs Default w60 guprawnienia3_grupy, Dodaj

Gui, uprawnienia3:Show, , Dodaj użytkownika
return

uprawnienia3_grupy:
Gui, uprawnienia3:Submit,Nohide
Gui, uprawnienia4:Submit,Nohide

If (grupa_upr_nazwa = "")
{
	MsgBox, 16, Ostrzeżenie, Uzupełnij nazwę użytkownika Windows
	return
}
upr_dir_uzytkownika = %Dane%\Uprawnienia\%grupa_upr_nazwa%.txt
FileDelete, %upr_dir_uzytkownika%
FileDelete, %grupa_0_upr%
If (grupa_upr_imie = "")
	grupa_upr_imie = %grupa_upr_nazwa%
FileAppend, %grupa_upr_4%`n%grupa_upr_imie%, %upr_dir_uzytkownika%

grupa_upr_0 := ""
GuiControl, uprawnienia2:, grupa_upr_0, |%grupa_upr_0%
grupa_upr_1 := ""
GuiControl, uprawnienia2:, grupa_upr_1, |%grupa_upr_1%
grupa_upr_2 := ""
GuiControl, uprawnienia2:, grupa_upr_2, |%grupa_upr_2%
grupa_upr_3 := ""
GuiControl, uprawnienia2:, grupa_upr_3, |%grupa_upr_3%

Loop, files, %Dane%\Uprawnienia\*, R
{ 
	SplitPath, A_LoopFileFullPath,,,,grupa_upr_0
	GuiControl, uprawnienia2:, grupa_upr_0, %grupa_upr_0%
	
	FileReadLine, grupa_upr_spr, %A_LoopFileFullPath%, 1
	If (grupa_upr_spr = "Pełne")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_1
		GuiControl, uprawnienia2:, grupa_upr_1, %grupa_upr_1%
	}
	If (grupa_upr_spr = "Rozszerzone")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_2
		GuiControl, uprawnienia2:, grupa_upr_2, %grupa_upr_2%
	}
	If (grupa_upr_spr = "Ograniczone")
	{
		SplitPath, A_LoopFileFullPath,,,,grupa_upr_3
		GuiControl, uprawnienia2:, grupa_upr_3, %grupa_upr_3%
	}
} 
grupa_upr_1 := ""
grupa_upr_2 := ""
grupa_upr_3 := ""
grupa_upr_0 := ""

grupa_upr_imie := ""
grupa_upr_nazwa := ""

Gui, uprawnienia3:Destroy
Gui, uprawnienia4:Destroy
return

info_program:
Gui, program:Destroy

Gui, program:Font, Bold
Gui, program:Add, Text, x10 y10 h20 w270 Section, Nazwa programu:
Gui, program:Font,
Gui, program:Add, Edit, y+5 Center ReadOnly w270 h20, %info_nazwa_programu%
Gui, program:Font, Bold
Gui, program:Add, Text, y+10 h20 w270, Wersja programu:
Gui, program:Font,
Gui, program:Add, Edit, y+5 Center ReadOnly w270 h20, %info_wersja_programu%
Gui, program:Font, Bold
Gui, program:Add, Text, y+10 h20 w270, Twórca programu:
Gui, program:Font,
Gui, program:Add, Edit, y+5 Center ReadOnly w270 h20, Grzegorz Czajka
Gui, program:Font, Bold
Gui, program:Add, Text, y+10 xs h20 w200, Uwagi i propozycje rozwoju:
Gui, program:Font,
Gui, program:Add, Button, x+10 w60 Default ginfo2_program, Prześlij
Gui, program:Font, Bold
Gui, program:Add, Text, y+10 xs h20 w200, Instrukcja programu:
Gui, program:Font,
Gui, program:Add, Button, x+10 w60 ginfo3_program, Przejdź

Gui, program:Show,,Informacje o programie
return

info2_program:
Run, https://edm.org.pl/?page_id=389
return
;try
;	outlookApp := ComObjActive("Outlook.Application")
;catch;
;	outlookApp := ComObjCreate("Outlook.Application")
;MailItem := outlookApp.CreateItem(0)
;MailItem.Recipients.Add("gczajka@gazoprojekt.pl")
;MailItem.Subject := "Uwagi i propozycje rozwoju"
;Tresc_maila = Poniżej zamieszczam swoje uwagi i propozycje rozwoju`n`n`nZ poważaniem`n%upr_imie% 
;MailItem.body := Tresc_maila
;MailItem.Display
;return

info3_program:
Run, https://edm.org.pl/?page_id=112
return

Nowy_pro:
Gui, npro:Destroy
Gui, Dok:Submit, noHide

inf_dok = Rodzaj dokumentacji:%A_Tab%`nNazwa Tomu/Obiektu:%A_Tab%`nNazwa Podtomu/części:%A_Tab%`nProjektant/Opracowujący:%A_Tab%`nSprawdzający:%A_Tab%%A_tab%`n

nr_dok := ""
Gui, npro:Font, Bold
Gui, npro:Add, Text, x10 y10 w200 Section, Numer dokument:
Gui, npro:Font,
Gui, npro:Add, Button, x+10 w60 gNowy_pro_gen, Generuj
Gui, npro:Add, Edit, y+5 xs w270 vnr_dok, %nr_dok%
Gui, npro:Font, Bold
Gui, npro:Add, Text, xs y+10 w200, Informacja o dokumencie:
Gui, npro:Font,
Gui, npro:Add, Button, x+10 w60 h20 gNowy_Pro_klucze, Klucze
Gui, npro:Add, Edit, y+5 xs w270 h200 vinf_dok, %inf_dok%
Gui, npro:Add, Button, y+10 w60 Default gnr_dok, Dodaj

Gui, npro:Show,, Tworzenie nowej dokumentacji
return

Nowy_Pro_klucze:
Gui, npro:Submit, noHide
Gui, Nowy_pro_klucze:Destroy

Gui, Nowy_pro_klucze:Font, Bold
Gui, Nowy_pro_klucze:Add, Text, x10 y10 w310 Section, Generuj opis dokumentacji na podstawie słów kluczowych:
Gui, Nowy_pro_klucze:Font,
Gui, Nowy_pro_klucze:Add, Text, y+15 w120, Rodzaj dokumentacji: 
Gui, Nowy_pro_klucze:Add, Edit, x+10 w180 vNEW_RODZAJ_DOKUMENTACJI, %Lista_rodz_dok%
Gui, Nowy_pro_klucze:Add, Text, y+5 xs+135 w120, NEW_RODZAJ_DOKUMENTACJI 
Gui, Nowy_pro_klucze:Add, Text, y+10 xs w120, Nazwa Tomu/Obiektu:
Gui, Nowy_pro_klucze:Add, Edit, x+10 w180 vNEW_NAZWA_TOMU_OBIEKTU, 
Gui, Nowy_pro_klucze:Add, Text, y+5 xs+135 w120, NEW_NAZWA_TOMU_OBIEKTU
Gui, Nowy_pro_klucze:Add, Text, y+10 xs w120, Nazwa Podtomu/części:
Gui, Nowy_pro_klucze:Add, Edit, x+10 w180 vNEW_NAZWA_PODTOMU_CZESCI, 
Gui, Nowy_pro_klucze:Add, Text, y+5 xs+135 w120, NEW_NAZWA_PODTOMU_CZESCI
Gui, Nowy_pro_klucze:Add, Text, y+10 xs w120, Projektant/Opracowujący:
Gui, Nowy_pro_klucze:Add, Edit, x+10 w180 vNEW_PROJEKTANT, 
Gui, Nowy_pro_klucze:Add, Text, y+5 xs+135 w120, NEW_PROJEKTANT
Gui, Nowy_pro_klucze:Add, Text, y+10 xs w120, Sprawdzający:
Gui, Nowy_pro_klucze:Add, Edit, x+10 w180 vNEW_SPRAWDZAJACY, 
Gui, Nowy_pro_klucze:Add, Text, y+5 xs+135 w120, NEW_SPRAWDZAJACY
Gui, Nowy_pro_klucze:Add, Button, y+10 xs Default w60 gNowy_Pro2_klucze, Generuj

Gui, Nowy_pro_klucze:Show,, Generuj klucze dokumentacji
return

Nowy_Pro2_klucze:
Gui, Nowy_pro_klucze:Submit,Nohide

inf_dok = Rodzaj dokumentacji:%A_Tab%%NEW_RODZAJ_DOKUMENTACJI%`nNazwa Tomu/Obiektu:%A_Tab%%NEW_NAZWA_TOMU_OBIEKTU%`nNazwa Podtomu/części:%A_Tab%%NEW_NAZWA_PODTOMU_CZESCI%`nProjektant/Opracowujący:%A_Tab%%NEW_PROJEKTANT%`nSprawdzający:%A_Tab%%A_tab%%NEW_SPRAWDZAJACY%`n

GuiControl, npro: , inf_dok, %inf_dok% 

Gui, Nowy_pro_klucze:Destroy
return

Nowy_pro_gen:
Gui, npro:Submit, noHide
Gui, Gen:Destroy

Numer_dokumetu := ""
branzadok := ""
rodzajdok := ""

Licznik_dir = %Dnae%\Licznik
IF !FileExist(Licznik_dir)
{
	FileCreateDir, %Licznik_dir%
	FileAppend, 00`n001`n-, %Licznik_dir%\info.txt
}
FileReadLine, Licznik1, %Licznik_dir%\info.txt, 1
FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2
FileReadLine, Licznik3, %Licznik_dir%\info.txt, 3

FileReadLine, NEW_NUMER_PROJEKTU, %Repozytorium%\info.txt, 1
SPR_NUMER_PROJEKTU := SubStr(NEW_NUMER_PROJEKTU,1,17) 
SPR2_NUMER_PROJEKTU := "Numer projektu:"A_Tab A_Tab
If (SPR_NUMER_PROJEKTU = SPR2_NUMER_PROJEKTU)
{
			NEW_NUMER_PROJEKTU := SubStr(NEW_NUMER_PROJEKTU, 18) 
			If (NEW_NUMER_PROJEKTU = "")
				NEW_NUMER_PROJEKTU := " "
}
else
	NEW_NUMER_PROJEKTU := ""
Lista_rodz_dok := ""

Gui, Gen:Add, Edit, x10 y10 w20 h20 Center ReadOnly, 1
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, 2
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, 3
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, 4
Gui, Gen:Add, Text, x+5 w10 h20 Center, %Licznik3%
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, P
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, B
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, 0
Gui, Gen:Add, Text, x+5 w10 h20 Center, %Licznik3%
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, 0
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, 1
Gui, Gen:Add, Text, x+5 w10 h20 Center, %Licznik3%
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, 0
Gui, Gen:Add, Edit, x+5 w20 h20 Center ReadOnly, 2
Gui, Gen:Add, Text, y+10 xs w95 h20 Center, Numer projektu
Gui, Gen:Add, Text, x+5 w95 h20 Center, Rodzaj dokumentu
Gui, Gen:Add, Text, x+5 w60 h20 Center, Tom/Obiekt
Gui, Gen:Add, Text, x+5 w50 h20 Center, Podtom
Gui, Gen:Add, Edit, y+5 xs w95 h20 Limit4 Center vnr_pr_gen, %NEW_NUMER_PROJEKTU%
Gui, Gen:Add, Text, x+5 w10 h20 Center, %Licznik3%
Gui, Gen:Add, Edit, x+5 w45 h20 Limit3 vLista_rodz_skr Center,
Gui, Gen:Add, Button, x+5 w20 h20 gR_dok Center, ...
Gui, Gen:Add, Text, x+5 w10 h20 Center, %Licznik3%
Gui, Gen:Add, Edit, x+5 w45 h20 Limit2 Number Center vtom_gen, 
Gui, Gen:Add, Text, x+5 w10 h20 Center, %Licznik3%
Gui, Gen:Add, Edit, x+5 w45 h20 Limit2 Number Center vpodtom_gen, 
Gui, Gen:Add, Button, y+10 xs w60 Default grodzaj_dok_gen, Generuj 

Gui, Gen:Show,, Generuj nazwę dokumentacji
return

rodzaj_dok_gen:
Gui, Gen:Submit, noHide
If (nr_pr_gen = "")
{
	MsgBox, 16, Ostrzeżenie, Numer projektu jest pusty
	return
}
If (Lista_rodz_skr = "")
{
	MsgBox, 16, Ostrzeżenie, Oznaczenie rodzaju dokumentacji jest pusty
	return
}
If (tom_gen = "")
{
	MsgBox, 16, Ostrzeżenie, Numer tomu jest pusty
	return
}
If (podtom_gen = "")
{
	MsgBox, 16, Ostrzeżenie, Numer podtomu jest pusty
	return
}
nr_dok = %nr_pr_gen%%Licznik3%%Lista_rodz_skr%%Licznik3%%tom_gen%%Licznik3%%podtom_gen%
GuiControl, npro: , nr_dok, %nr_dok%

Gui, Gen:Destroy
return

R_dok:
Gui, R_gen:Destroy

klasa := "Podstawowe"
Rodzaj_dok_dr = %Dane%\Rodzaje_dok
If !FileExist(Rodzaj_dok_dr)
	FileCreateDir, %Dane%\Rodzaje_dok
Gui, R_gen:Font, Bold
Gui, R_gen:Add, Text, x10 y10 w130, Rodzaje dokumentacji:
Gui, R_gen:Font,
Gui, R_gen:Add, Button, x+10 w60 gdodaj_rodz_dok, Dodaj
Gui, R_gen:Add, Button, x+10 w60 gedytuj_rodz_dok, Edytuj
Gui, R_gen:Add, Button, x+10 w60 gusun_dok_rodz, Usuń
Gui, R_gen:Add, DropDownList, y+5 xs w340 gklasa_rodz vklasa, Podstawowe||Rozszerzone
Gui, R_gen:Add, ListBox, y+10 w340 h200 vLista_rodz_dok gLista_rodz_dok, %Lista_rodz_dok%
Gui, R_gen:Add, Edit, y+10 w340 Center ReadOnly vLista_rodz_skr, %Lista_rodz_skr%
Gui, R_gen:Add, Button, y+10 w60 Default gwstaw_rodz, Wstaw

Lista_rodz_dok := "Wybierz rodzaj dokumentacji||"
GuiControl, R_gen: , Lista_rodz_dok, |%Lista_rodz_dok%

Loop, files, %Dane%\Rodzaje_dok\*, R
{ 
	FileReadLine, klasa2, %A_LoopFileFullPath%, 2
	If (klasa2 = klasa)
	{
		SplitPath, A_LoopFileFullPath,,,,Lista_rodz_dok
		dok_2_dir = %A_LoopFileFullPath%
		GuiControl, R_gen: , Lista_rodz_dok, %Lista_rodz_dok% 				
	}
} 

Gui, R_gen:Show,, Rodzaj dokumentacji
return

dodaj_rodz_dok:
Gui, R_gen:Submit, Nohide
Gui, R2_gen:Destroy

klasa3 = Podstawowe

Gui, R2_gen:Add, Text, y10 x10 h20 w200, Nazwa rodzaju dokumentacji:
Gui, R2_gen:Add, Edit, y+5 w200 vdodaj_nazwa_rodz,
Gui, R2_gen:Add, Text, y+10 h20 w200, Skrót rodzaju dokumentacji:
Gui, R2_gen:Add, Edit, y+5 w200 limit3 vdodaj_skr_rodz,
Gui, R2_gen:Add, Text, y+10 h20 w200, Klasyfikacja:
Gui, R2_gen:Add, DropDownList, y+5 w200 vklasa3, Podstawowe||Rozszerzone
Gui, R2_gen:Add, Button, y+10 w60 Default gdodaj_rodz_2, Dodaj

Gui, R2_gen:Show,, Dodaj rodzaj dokumentacji
return

edytuj_rodz_dok:
Gui, R_gen:Submit, Nohide
Gui, R3_gen:Destroy

If (Lista_rodz_dok = "Wybierz rodzaj dokumentacji")
{
	MsgBox, 16, Ostrzeżenie, Wybierz rodzaj dokumentacji
	return
}
klasa3 = %klasa%
If (klasa = "Podstawowe")
{
	opcja_r_1 = Podstawowe
	opcja_r_2 = Rozszerzone
}
If (klasa = "Rozszerzone")
{
	opcja_r_2 = Podstawowe
	opcja_r_1 = Rozszerzone
}

Gui, R3_gen:Add, Text, y10 x10 h20 w200, Nazwa rodzaju dokumentacji:
Gui, R3_gen:Add, Edit, y+5 w200 vdodaj_nazwa_rodz, %Lista_rodz_dok% 
Gui, R3_gen:Add, Text, y+10 h20 w200, Skrót rodzaju dokumentacji:
Gui, R3_gen:Add, Edit, y+5 w200 limit3 vdodaj_skr_rodz, %Lista_rodz_skr%
Gui, R3_gen:Add, Text, y+10 h20 w200, Klasyfikacja:
Gui, R3_gen:Add, DropDownList, y+5 w200 vklasa3, %opcja_r_1%||%opcja_r_2%
Gui, R3_gen:Add, Button, y+10 w60 Default gdodaj_rodz_2, Edytuj

Gui, R3_gen:Show,, Edytuj rodzaj dokumentacji
return

usun_dok_rodz:
Gui, Dok:Submit, noHide
Gui, R4_gen:Destroy

If (Lista_rodz_dok = "Wybierz rodzaj dokumentacji")
{
	MsgBox, 16, Ostrzeżenie, Wybierz rodzaj dokumentacji
	return
}

Gui, R4_gen:Add, Text, x10 y10 h30 w290, Czy na pewno chcesz usunąć zaznaczony rodzaj dokumentacji?
Gui, R4_gen:Add, Button, y+10 xp+80 w60 gUsun_rodz_2, Usuń
Gui, R4_gen:Add, Button, x+10 w60 Default gUsun_rodz_3 , Anuluj

Gui, R4_gen:Show,, Ostrzeżenie!
return

Usun_rodz_2:
Gui, R4_gen:Submit, noHide

FileDelete, %Dane%\Rodzaje_dok\%Lista_rodz_dok%.txt

Lista_rodz_dok := "Wybierz rodzaj dokumentacji||"
GuiControl, R_gen: , Lista_rodz_dok, |%Lista_rodz_dok%

Loop, files, %Dane%\Rodzaje_dok\*, R
{ 
	FileReadLine, klasa2, %A_LoopFileFullPath%, 2
	If (klasa2 = klasa)
	{
		SplitPath, A_LoopFileFullPath,,,,Lista_rodz_dok
		dok_2_dir = %A_LoopFileFullPath%
		GuiControl, R_gen: , Lista_rodz_dok, %Lista_rodz_dok% 				
	}
} 

Lista_rodz_skr := ""
GuiControl, R_gen: , Lista_rodz_skr, %Lista_rodz_skr%

Gui, R4_gen:Destroy
return

Usun_rodz_3:
Gui, R4_gen:Destroy
return

dodaj_rodz_2:
Gui, R2_gen:Submit,
Gui, R3_gen:Submit,

If (dodaj_nazwa_rodz = "")
{
	MsgBox, 16, Ostrzeżenie, Nazwa rodzaju dokumentacji jest pusta
	return
}

FileDelete, %Dane%\Rodzaje_dok\%dodaj_nazwa_rodz%.txt
FileAppend, %dodaj_skr_rodz%`n%klasa3%, %Dane%\Rodzaje_dok\%dodaj_nazwa_rodz%.txt

Lista_rodz_dok := "Wybierz rodzaj dokumentacji||"
GuiControl, R_gen: , Lista_rodz_dok, |%Lista_rodz_dok%

Loop, files, %Dane%\Rodzaje_dok\*, R
{ 
	FileReadLine, klasa2, %A_LoopFileFullPath%, 2
	If (klasa2 = klasa)
	{
		SplitPath, A_LoopFileFullPath,,,,Lista_rodz_dok
		dok_2_dir = %A_LoopFileFullPath%
		GuiControl, R_gen: , Lista_rodz_dok, %Lista_rodz_dok% 				
	}
} 

Lista_rodz_skr := ""
GuiControl, R_gen: , Lista_rodz_skr, %Lista_rodz_skr%

Gui, R2_gen:Destroy
Gui, R3_gen:Destroy
return

klasa_rodz:
Gui, R_gen:Submit, Nohide

Lista_rodz_dok := "Wybierz rodzaj dokumentacji||"
GuiControl, R_gen: , Lista_rodz_dok, |%Lista_rodz_dok%

Loop, files, %Dane%\Rodzaje_dok\*, R
{ 
	FileReadLine, klasa2, %A_LoopFileFullPath%, 2
	If (klasa2 = klasa)
	{
		SplitPath, A_LoopFileFullPath,,,,Lista_rodz_dok
		dok_2_dir = %A_LoopFileFullPath%
		GuiControl, R_gen: , Lista_rodz_dok, %Lista_rodz_dok% 				
	}
} 

Lista_rodz_skr := ""
GuiControl, R_gen: , Lista_rodz_skr, %Lista_rodz_skr%

return

Lista_rodz_dok:
Gui, R_gen:Submit, Nohide

FileReadLine, Lista_rodz_skr, %Dane%\Rodzaje_dok\%Lista_rodz_dok%.txt, 1
GuiControl, R_gen: , Lista_rodz_skr, %Lista_rodz_skr%

return

wstaw_rodz:
Gui, R_gen:Submit, Nohide

If (Numer_dok_2 = "Wybierz rodzaj dokumentacji")
{
	MsgBox, 16, Ostrzeżenie, Wybierz rodzaj dokumentacji
	return
}

GuiControl, Gen: , Lista_rodz_skr, %Lista_rodz_skr%

inf_dok = Rodzaj dokumentacji:%A_Tab%%Lista_rodz_dok%`nNazwa Tomu/Obiektu:%A_Tab%`nNazwa Podtomu/części:%A_Tab%`nProjektant/Opracowujący:%A_Tab%`nSprawdzający:%A_Tab%%A_tab%`n

GuiControl, npro: , inf_dok, %inf_dok%

Gui, R_gen:Destroy
return

nr_dok:
Gui, npro:Submit, noHide

znaki_spec_spr = %nr_dok_2%
Gosub, znaki_spec
If (znak_test = 1)
	return
nr_dok_dr = %Repozytorium%\%nr_dok%
If FileExist(nr_dok_dr)
{
	MsgBox, 16, Ostrzeżenie!, Dokumentacja już istnieje
	return
}
FileCreateDir, %Repozytorium%\%nr_dok%
FileAppend, %inf_dok%, %Repozytorium%\%nr_dok%\info.txt
FileCreateDir, %Roboczy%\%nr_dok%

Numer_dok_2 :="Wybierz dokumentację||"
GuiControl, dok: , Numer_dok_2, |%Numer_dok_2% 
Loop, files, %Repozytorium%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	GuiControl, dok: , Numer_dok_2, %Numer_dok_2% 
} 
Numer_dok_2 := "Wybierz dokumentację"
dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%
Gosub, radio_filtr_clean
dok_2 := "Wybierz dokument"
Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, |%Opis_dok% 

Gui, npro:Destroy
return

Edytuj_pro:
Gui, edpro:Destroy
Gui, Dok:Submit, noHide

If (Numer_dok_2 = "Wybierz dokumentację")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokumentację
	Gui, edpro:Destroy
	return
}

zm_nr_dok := ""
Gui, edpro:Font, Bold
Gui, edpro:Add, Text, x10 y10 Section, Numer dokument:
Gui, edpro:Font,
Gui, edpro:Add, Edit, y+10 vzm_nr_dok w270, %Numer_dok_2%
Gui, edpro:Add, Checkbox, y+10 vchecked_projekt_klucze w270, Zmień wewnątrz dokumentów „docx”
Gui, edpro:Add, Button, xs y+10 Default gzm_nr_dok, Zmień

Gui, edpro:Show,, Edycja nazwy
return

zm_nr_dok:
Gui, edpro:Submit, noHide

SplashTextOn, 180,40, Proces, Trwa proces...
FileMoveDir, %Roboczy%\%Numer_dok_2% ,%Roboczy%\%zm_nr_dok%, 2
FileMoveDir, %Repozytorium%\%Numer_dok_2% ,%Repozytorium%\%zm_nr_dok%, 2
Replace = %Numer_dok_2%
With = %zm_nr_dok%
Numer_dok_2 = %zm_nr_dok%

Loop Files, %Roboczy%\%Numer_dok_2%\*, R
{
	new_name := StrReplace(A_LoopFileName, Replace, With)
	new_name = %Roboczy%\%Numer_dok_2%\%new_name%
	FileMove,%A_LoopFileFullPath%,%new_name%, 1
}
Loop Files, %Repozytorium%\%Numer_dok_2%\*, D
{
	new_name := StrReplace(A_LoopFileName, Replace, With)
	new_name_path = %Repozytorium%\%Numer_dok_2%\%new_name%
	FileMoveDir,%A_LoopFileFullPath%,%new_name_path%, 2
	
	FileReadLine, Org_dok_2, %new_name_path%\info.txt, 1
	FileReadLine, Nr_dok_2, %new_name_path%\info.txt, 2
	FileReadLine, Rewizja_1, %new_name_path%\info.txt, 3
	FileReadLine, Rewizja_2, %new_name_path%\info.txt, 4
	FileReadLine, Rewizja_4, %new_name_path%\info.txt, 5
	FileReadLine, Status_dok, %new_name_path%\info.txt, 6
	FileReadLine, dok_ext, %new_name_path%\info.txt, 7
	FileRead, His_dok, %new_name_path%\his.txt
	Rewizja_3 := Format("{01:02}", Rewizja_2)	
	Rewizja_4 += 1
	Rewizja_5 := Format("{01:03}", Rewizja_4)
	FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
	
	OLD_NUMER_DOKUMENTU = %A_LoopFileName%
	
	Org_dok_2 = %new_name%.%dok_ext%
	Nr_dok_2 = %new_name%
	
	FileDelete, %new_name_path%\info.txt
	FileAppend, %Org_dok_2%`n%Nr_dok_2%`n%Rewizja_1%`n%Rewizja_3%`n%Rewizja_5%`n%Status_dok%`n%dok_ext%`n, %new_name_path%\info.txt
	
	FileDelete, %new_name_path%\his.txt
	FileAppend, %Rewizja_1%.%Rewizja_3%.%Rewizja_5%`n%data%`nZmieniono nazwę pliku z %A_LoopFileName% na %new_name% przez %upr_imie%`n`n%his_dok%, %new_name_path%\his.txt
	
	If (dok_ext = "docx" and checked_projekt_klucze = 1)
	{
		NEW_NUMER_DOKUMENTU = %new_name%
		Path = %Roboczy%\%Numer_dok_2%\%Org_dok_2%
		oWord := ComObjCreate("Word.Application")
		oWord.Visible := false
		oWord.Documents.Open(Path)
		FindAndReplace( oWord, OLD_NUMER_DOKUMENTU, NEW_NUMER_DOKUMENTU )
		oWord.Documents.Save()
		oWord.quit()	
	}	
}

SplashTextOff

Numer_dok_2 := "Wybierz dokumentację||"
GuiControl, dok: , Numer_dok_2, |%Numer_dok_2% 
Loop, files, %Repozytorium%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	GuiControl, dok: , Numer_dok_2, %Numer_dok_2% 
} 
Numer_dok_2 := "Wybierz dokumentację"
dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%
Gosub, radio_filtr_clean
dok_2 := "Wybierz dokument"
Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

Gui, edpro:Destroy
return

Usun_pro:
Gui, Dok:Submit, noHide
Gui, usun_pro:Destroy

If (Numer_dok_2 = "Wybierz dokumentację")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokumentację
	return
}

Gui, usun_pro:Add, Text, x10 y10 h20 w290, Czy na pewno chcesz usunąć zaznaczoną dokumentację?
Gui, usun_pro:Add, Button, y+10 x80 w60 gUsun_pro_2, Usuń
Gui, usun_pro:Add, Button, x+10 w60 Default gUsun_pro_3 , Anuluj

Gui, usun_pro:Show,, Ostrzeżenie!
return

Usun_pro_2:
SplashTextOn, 180,40, Proces, Trwa proces...
FileRemoveDir, %Roboczy%\%Numer_dok_2%, 1
FileRemoveDir, %Repozytorium%\%Numer_dok_2%, 1
SplashTextOff
Numer_dok_2 :="Wybierz dokumentację||"
GuiControl, dok: , Numer_dok_2, |%Numer_dok_2% 
Loop, files, %Repozytorium%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	GuiControl, dok: , Numer_dok_2, %Numer_dok_2% 
} 
Numer_dok_2 := "Wybierz dokumentację"
dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%
Gosub, radio_filtr_clean
dok_2 := "Wybierz dokument"
Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

Gui, usun_pro:Destroy
return

Usun_pro_3:
Gui, usun_pro:Destroy
return

Wydaj_pro:
Gui, Dok:Submit, noHide

If (Numer_dok_2 = "Wybierz dokumentację")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokumentację
	return
}

Loop, files, %Repozytorium%\%Numer_dok_2%\*, D
{ 
	FileReadLine, Status_dok, %A_LoopFileFullPath%\info.txt, 6
	IF (Status_dok = "Roboczy" or Status_dok = "Sprawdzany" or Status_dok = "Odrzucony")
	{
		MsgBox, 16, Ostrzeżenie, Aby wydać dokumentację wszystkie dokumenty muszą uzyskać status: Akceptacja
		return
	}
} 

Edytowalne_dir = %Dokumentacja%\Edytowalne
Nieedytowalne_dir = %Dokumentacja%\Nieedytowalne
If !FileExist(Edytowalne_dir)
	FileCreateDir, %Dokumentacja%\Edytowalne
If !FileExist(Nieedytowalne_dir)
	FileCreateDir, %Dokumentacja%\Nieedytowalne
Dok_ed_dr = %Edytowalne_dir%\%Numer_dok_2%
Dok_ned_dr = %Nieedytowalne_dir%\%Numer_dok_2%

FileRemoveDir, %Dok_ed_dr%, 1
FileCreateDir, %Dok_ed_dr%
FileRemoveDir, %Dok_ned_dr%, 1
FileCreateDir, %Dok_ned_dr%

FormatTime, data_2, A_Now, yyyy_MM_dd_HHmmss
FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
SplashTextOn, 180,40, Proces, Trwa proces...

Loop, files, %Repozytorium%\%Numer_dok_2%\*, D
{ 
	FileReadLine,Org_dok_2, %A_LoopFileFullPath%\info.txt, 1
	FileReadLine,Nr_dok_2, %A_LoopFileFullPath%\info.txt, 2
	FileReadLine,Rewizja_1, %A_LoopFileFullPath%\info.txt, 3
	FileReadLine,Rewizja_2, %A_LoopFileFullPath%\info.txt, 4
	FileReadLine,dok_ext, %A_LoopFileFullPath%\info.txt, 7
	FileRead, his_dok, %Repozytorium%\%Numer_dok_2%\%Nr_dok_2%\his.txt
	Rewizja_3 := Format("{01:02}", Rewizja_2)	
	FileCopy, %Roboczy%\%Numer_dok_2%\%Org_dok_2%, %Dok_ed_dr%\%Numer_dok%\%Nr_dok_2%-R%Rewizja_1%_%Rewizja_3%.%dok_ext%, 1
	
	FileCreateDir, %Repozytorium%\%Numer_dok_2%\%Nr_dok_2%\%data_2%_R%Rewizja_1%_%Rewizja_3%_Wyd
	FileCopy, %Roboczy%\%Numer_dok_2%\%Org_dok_2%, %Repozytorium%\%Numer_dok_2%\%Nr_dok_2%\%data_2%_R%Rewizja_1%_%Rewizja_3%_Wyd\%Org_dok_2%, 1
	FileSetAttrib, +R, %Repozytorium%\%Numer_dok_2%\%Nr_dok_2%\%data_2%_R%Rewizja_1%_%Rewizja_3%_Wyd\%Org_dok_2%
	
	Rewizja_4++
	Rewizja_5 := Format("{01:03}", Rewizja_4)
	
	FileDelete, %Repozytorium%\%Numer_dok_2%\%Nr_dok_2%\Info.txt
	FileAppend, %Org_dok_2%`n%Nr_dok_2%`n%Rewizja_1%`n%Rewizja_3%`n%Rewizja_5%`nWydany`n%dok_ext%`n, %Repozytorium%\%Numer_dok_2%\%Nr_dok_2%\info.txt
	
	FileDelete, %Repozytorium%\%Numer_dok_2%\%Nr_dok_2%\his.txt
	FileAppend, %Rewizja_1%.%Rewizja_3%.%Rewizja_5%`n%data%`nDokument wydany w wersji %Rewizja_1%_%Rewizja_3% oraz zarchiwizowana przez %upr_imie%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%Nr_dok_2%\his.txt
} 
Loop, files, %Repozytorium%\%Numer_dok_2%\*.pdf, R
{ 
	SplitPath, A_LoopFileFullPath, Numer_dok_pdf
	FileCopy, %A_LoopFileFullPath%, %Dok_ned_dr%\%Numer_dok_pdf%, 1
} 
Loop, files, %Repozytorium%\%Numer_dok_2%\*.rar, R
{ 
	SplitPath, A_LoopFileFullPath, Numer_dok_rar
	FileCopy, %A_LoopFileFullPath%, %Dok_ed_dr%\%Numer_dok_rar%, 1
} 
Loop, files, %Repozytorium%\%Numer_dok_2%\*.zip, R
{ 
	SplitPath, A_LoopFileFullPath, Numer_dok_zip
	FileCopy, %A_LoopFileFullPath%, %Dok_ed_dr%\%Numer_dok_zip%, 1
} 
Loop, files, %Repozytorium%\%Numer_dok_2%\*.7z, R
{ 
	SplitPath, A_LoopFileFullPath, Numer_dok_7z
	FileCopy, %A_LoopFileFullPath%, %Dok_ed_dr%\%Numer_dok_7z%, 1
} 
SplashTextOff

dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%

Loop, files, %Repozytorium%\%Numer_dok_2%\*%filtr2%*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,dok_2
	FileReadLine, Spr_stat, %A_LoopFileFullPath%\info.txt, 6
	If (filtr3 = "")
		GuiControl, dok: , dok_2, %dok_2%
	If else (Filtr3 = Spr_stat)
		GuiControl, dok: , dok_2, %dok_2%
} 
dok_2 := "Wybierz dokument"

Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 


Status_pro := ""
Status1_pro := ""
Status2_pro := ""
Status3_pro := ""

GuiControl, dok: , Status_pro, %Status_pro%

Status_pro := "Wydana"
GuiControl, dok: , Status_pro, %Status_pro%	

MsgBox, 64, Informacja, Dokumentacja została wydana
return

Odswiez:
Gui, dok:Submit, noHide

Numer_dok_2 :="Wybierz dokumentację||"
GuiControl, dok: , Numer_dok_2, |%Numer_dok_2% 
Loop, files, %Repozytorium%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	GuiControl, dok: , Numer_dok_2, %Numer_dok_2% 
} 
Numer_dok_2 := "Wybierz dokumentację"
dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%
dok_2 := "Wybierz dokument"
Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

return

Przelacz:
Gui, dok:Submit, noHide
Gui, przelacz:Destroy

Gui, przelacz:Font, Bold
Gui, przelacz:Add, Text, x10 y10 h20 Section w340, Folder z projektami:
Gui, przelacz:Font,
Gui, przelacz:Add, Edit, y+5 w340 h40 ReadOnly vdir_projekty, %dir_projekty%
Gui, przelacz:Add, Button, y+5 w60 gzmien_dr_pr, Zmień
Gui, przelacz:Font, Bold
Gui, przelacz:Add, Text, y+10 h20 w340, Dostępne projekty:
Gui, przelacz:Font,
Gui, przelacz:Add, ListBox, y+5 xs w340 h200 vprojekt_dir, %projekt_dir%
Gui, przelacz:Add, Button, y+5 w60 Default gprzelacz_dr_pr, Przełącz
Gui, przelacz:Add, Button, x+10 w60 vButtonToHide1 gedytuj_dr_pr, Edytuj
Gui, przelacz:Add, Button, x+10 w60 vButtonToHide2 gdodaj_dr_pr, Dodaj
Gui, przelacz:Add, Button, x+10 w60 vButtonToMove1 gzamknij_dr_pr, Zamknij	

If else (upr_uzytkownika = "Rozszerzone" or upr_uzytkownika = "Ograniczone")
{
	GuiControl,przelacz:Hide,ButtonToHide1
	GuiControl,przelacz:Hide,ButtonToHide2
	GuiControl,przelacz:Move,ButtonToMove1, x80
}	

projekt_dir := "Wybierz projekt||"
GuiControl, przelacz: , projekt_dir, |%projekt_dir% 
Loop, files, %dir_projekty%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,projekt_dir
	projekt_dir_2 = %A_LoopFileFullPath%\info_DMS.txt
	IF FileExist(projekt_dir_2)
		GuiControl, przelacz: , projekt_dir, %projekt_dir% 
} 

Gui, przelacz:Show,, Przełącz projekt
return

przelacz_dr_pr:
Gui, przelacz:Submit, noHide

If (projekt_dir = "Wybierz projekt")
{
	MsgBox, 16, Ostrzeżenie, Wybierz projekt
	return
}

FileReadLine, Dane, %dir_projekty%\%projekt_dir%\Info_DMS.txt, 1
FileReadLine, Roboczy, %dir_projekty%\%projekt_dir%\Info_DMS.txt, 2
FileReadLine, Wzor, %dir_projekty%\%projekt_dir%\Info_DMS.txt, 3
FileReadLine, Repozytorium, %dir_projekty%\%projekt_dir%\Info_DMS.txt, 4
FileReadLine, Dokumentacja, %dir_projekty%\%projekt_dir%\Info_DMS.txt, 5
FileReadLine, dir_projekty, %dir_projekty%\%projekt_dir%\Info_DMS.txt, 6

FileRead, Info_projekt, %Repozytorium%/info.txt
GuiControl, dok: , Info_projekt, %Info_projekt% 

Numer_dok_2 :="Wybierz dokumentację||"
GuiControl, dok: , Numer_dok_2, |%Numer_dok_2% 
Loop, files, %Repozytorium%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	GuiControl, dok: , Numer_dok_2, %Numer_dok_2% 
} 
Numer_dok_2 := "Wybierz dokumentację"
dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%
dok_2 := "Wybierz dokument"
Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

WinSetTitle, %Tytul_projektu%, , %projekt_dir%
Tytul_projektu = %projekt_dir%

return

zmien_dr_pr:
Gui, przelacz:Submit, noHide
FileSelectFolder, dir_projekty
GuiControl, przelacz: , dir_projekty, %dir_projekty% 	

projekt_dir := "Wybierz projekt||"
GuiControl, przelacz: , projekt_dir, |%projekt_dir% 
Loop, files, %dir_projekty%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,projekt_dir
	projekt_dir_2 = %A_LoopFileFullPath%\info_DMS.txt
	IF FileExist(projekt_dir_2)
		GuiControl, przelacz: , projekt_dir, %projekt_dir% 
}
return

edytuj_dr_pr:
Gui, dodaj2:Destroy
Gui, przelacz:Submit, noHide

If (projekt_dir = "Wybierz projekt")
{
	MsgBox, 16, Ostrzeżenie, Wybierz projekt
	return
}

SplitPath, A_ScriptDir,,,,Nazwa_projektu

FileReadLine, Dane, %A_ScriptDir%\Info_DMS.txt, 1
FileReadLine, Roboczy, %A_ScriptDir%\Info_DMS.txt, 2
FileReadLine, Wzor, %A_ScriptDir%\Info_DMS.txt, 3
FileReadLine, Repozytorium, %A_ScriptDir%\Info_DMS.txt, 4
FileReadLine, Dokumentacja, %A_ScriptDir%\Info_DMS.txt, 5
FileReadLine, dir_projekty, %A_ScriptDir%\Info_DMS.txt, 6
Roboczy_pomocniczy = %Roboczy%
Repozytorium_pomocniczy = %Repozytorium%

Gui, dodaj2:Add, Text, x10 y10 w150 h20, Nazwa projektu:
Gui, dodaj2:Add, Edit, x+10 w400 vNazwa_projektu ReadOnly, %Nazwa_projektu%
Gui, dodaj2:Add, Text, xs y+10 h20 w150, Folder z danymi pomocniczymi:
Gui, dodaj2:Add, Edit, x+10 w400 vDane, %Dane%
Gui, dodaj2:Add, Text, xs y+10 h20 w150, Folder roboczy
Gui, dodaj2:Add, Edit, x+10 w400 vRoboczy, %Roboczy%
Gui, dodaj2:Add, Text, xs y+10 h20 w150, Folder z szablonami:
Gui, dodaj2:Add, Edit, x+10 w400 vWzor, %Wzor%
Gui, dodaj2:Add, Text, xs y+10 h20 w150, Folder rezpozytorium:
Gui, dodaj2:Add, Edit, x+10 w400 vRepozytorium, %Repozytorium%
Gui, dodaj2:Add, Text, xs y+10 h20 w150, Folder z dokumentacją:
Gui, dodaj2:Add, Edit, x+10 w400 vDokumentacja, %Dokumentacja%
Gui, dodaj2:Add, Checkbox, y+10 xs h20 w400 vukryc_DMS checked, Ukryj folder: Dane, Repozytoriu, Szablony
Gui, dodaj2:Add, Button, y+10 w60 Default gutworz2_dr_pr, Zmień

Gui, dodaj2:Show,, Edytuj projekt
return

dodaj_dr_pr:
Gui, dodaj:Destroy
Gui, przelacz:Submit, noHide

SplitPath, A_ScriptDir,,,,Nazwa_projektu2
FileReadLine, Dane2, %A_ScriptDir%\Info_DMS.txt, 1
FileReadLine, Roboczy2, %A_ScriptDir%\Info_DMS.txt, 2
FileReadLine, Wzor2, %A_ScriptDir%\Info_DMS.txt, 3
FileReadLine, Repozytorium2, %A_ScriptDir%\Info_DMS.txt, 4
FileReadLine, Dokumentacja2, %A_ScriptDir%\Info_DMS.txt, 5
FileReadLine, dir_projekty2, %A_ScriptDir%\Info_DMS.txt, 6
FileRead, Info_projekt2, %Repozytorium%/info.txt

Info_projekt2 = Numer projektu:%A_Tab%%A_Tab%`nNazwa inwestycji:%A_Tab%%A_Tab%`nAdres inwestycji:%A_Tab%%A_Tab%`nNazwa zamawiającego:%A_Tab%`nAdres zamawiającego:%A_Tab%`nNumer umowy:%A_Tab%%A_Tab%`nKierownik projektu:%A_Tab%%A_Tab%`n

Gui, dodaj:Add, Text, x10 y10 h20 w150, Nazwa projektu:
Gui, dodaj:Add, Edit, x+10 w400 vNazwa_projektu2 gNazwa_projektu2, %Nazwa_projektu2%
Gui, dodaj:Add, Text, xs y+10 w150, Folder z danymi pomocniczymi:
Gui, dodaj:Add, Edit, x+10 w400 vDane2, %Dane2%
Gui, dodaj:Add, Text, xs y+10 w150, Folder roboczy
Gui, dodaj:Add, Edit, x+10 w400 vRoboczy2, %Roboczy2%
Gui, dodaj:Add, Text, xs y+10 w150, Folder z szablonami:
Gui, dodaj:Add, Edit, x+10 w400 vWzor2, %Wzor2%
Gui, dodaj:Add, Text, xs y+10 w150, Folder rezpozytorium:
Gui, dodaj:Add, Edit, x+10 w400 vRepozytorium2, %Repozytorium2%
Gui, dodaj:Add, Text, xs y+10 w150, Folder z dokumentacją:
Gui, dodaj:Add, Edit, x+10 w400 vDokumentacja2, %Dokumentacja2%
Gui, dodaj:Add, Checkbox, y+10 xs h20 w400 vukryc_DMS checked, Ukryj folder: Dane, Repozytoriu, Szablony
Gui, dodaj:Add, Text, xs y+10 w490 h20, Informacje o projekcie:
Gui, dodaj:Add, Button, x+10 w60 h20 gPro_klucze, Klucze
Gui, dodaj:Add, Edit, y+5 xs w560 h200 vInfo_projekt2, %Info_projekt2%
Gui, dodaj:Add, Button, xs y+10 w60 Default gutworz_dr_pr, Utwórz

Gui, dodaj:Show,, Tworzenie nowego projektu
return

Pro_klucze:
Gui, przelacz:Submit, noHide
Gui, pro_klucze:Destroy

Gui, pro_klucze:Font, Bold
Gui, pro_klucze:Add, Text, x10 y10 w310 Section, Generuj opis projektu na podstawie słów kluczowych:
Gui, pro_klucze:Font,
Gui, pro_klucze:Add, Text, y+15 w120, Numer projektu:  
Gui, pro_klucze:Add, Edit, x+10 w180 vNEW_NUMER_PROJEKTU, 
Gui, pro_klucze:Add, Text, y+5 xs+135 w120, NEW_NUMER_PROJEKTU 
Gui, pro_klucze:Add, Text, y+10 xs w120, Nazwa inwestycji:
Gui, pro_klucze:Add, Edit, x+10 w180 vNEW_NAZWA_INWESTYCJI, 
Gui, pro_klucze:Add, Text, y+5 xs+135 w120, NEW_NAZWA_INWESTYCJI
Gui, pro_klucze:Add, Text, y+10 xs w120, Adres inwestycji:
Gui, pro_klucze:Add, Edit, x+10 w180 vNEW_ADRES_INWESTYCJI, 
Gui, pro_klucze:Add, Text, y+5 xs+135 w120, NEW_ADRES_INWESTYCJI
Gui, pro_klucze:Add, Text, y+10 xs w120, Nazwa zamawiającego
Gui, pro_klucze:Add, Edit, x+10 w180 vNEW_NAZWA_ZAMAWIAJĄCEGO, 
Gui, pro_klucze:Add, Text, y+5 xs+135 w120, NEW_NAZWA_ZAMAWIAJĄCEGO
Gui, pro_klucze:Add, Text, y+10 xs w120, Adres zamawiającego:
Gui, pro_klucze:Add, Edit, x+10 w180 vNEW_ADRES_ZAMAWIAJACEGO, 
Gui, pro_klucze:Add, Text, y+5 xs+135 w120, NEW_ADRES_ZAMAWIAJACEGO
Gui, pro_klucze:Add, Text, y+10 xs w120, Numer umowy:
Gui, pro_klucze:Add, Edit, x+10 w180 vNEW_NUMER_UMOWY, 
Gui, pro_klucze:Add, Text, y+5 xs+135 w120, NEW_NUMER_UMOWY
Gui, pro_klucze:Add, Text, y+10 xs w120, Kierownik projektu:
Gui, pro_klucze:Add, Edit, x+10 w180 vNEW_KIEROWNIK_PROJEKTU, 
Gui, pro_klucze:Add, Text, y+5 xs+135 w120, NEW_KIEROWNIK_PROJEKTU
Gui, pro_klucze:Add, Button, y+10 xs Default w60 gPro2_klucze, Generuj

Gui, pro_klucze:Show,, Generuj klucze projektu
return

Pro2_klucze:
Gui, pro_klucze:Submit,Nohide

Info_projekt2 = Numer projektu:%A_Tab%%A_Tab%%NEW_NUMER_PROJEKTU%`nNazwa inwestycji:%A_Tab%%A_Tab%%NEW_NAZWA_INWESTYCJI%`nAdres inwestycji:%A_Tab%%A_Tab%%NEW_ADRES_INWESTYCJI%`nNazwa zamawiającego:%A_Tab%%NEW_NAZWA_ZAMAWIAJĄCEGO%`nAdres zamawiającego:%A_Tab%%NEW_ADRES_ZAMAWIAJACEGO%`nNumer umowy:%A_Tab%%A_Tab%%NEW_NUMER_UMOWY%`nKierownik projektu:%A_Tab%%A_Tab%%NEW_KIEROWNIK_PROJEKTU%`n
GuiControl, dodaj: , Info_projekt2, %Info_projekt2% 

Gui, pro_klucze:Destroy
return

Nazwa_projektu2:
Gui, dodaj:Submit, noHide
Dane2 = %dir_projekty%\%Nazwa_projektu2%\Dane
GuiControl, dodaj: , Dane2, %Dane2% 
Roboczy2 = %dir_projekty%\%Nazwa_projektu2%\Roboczy
GuiControl, dodaj: , Roboczy2, %Roboczy2% 
Wzor2 = %dir_projekty%\%Nazwa_projektu2%\Wzor
GuiControl, dodaj: , Wzor2, %Wzor2% 
Repozytorium2 = %dir_projekty%\%Nazwa_projektu2%\Rep
GuiControl, dodaj: , Repozytorium2, %Repozytorium2% 
Dokumentacja2 = %dir_projekty%\%Nazwa_projektu2%\Dokumentacja
GuiControl, dodaj: , Dokumentacja2, %Dokumentacja2% 
return

utworz_dr_pr:
Gui, dodaj:Submit, noHide

SplashTextOn, 180,40, Proces, Trwa proces...

nowy_dodaj_dr = %dir_projekty%\%Nazwa_projektu2%
If !FileExist(nowy_dodaj_dr)
	FileCreateDir, %dir_projekty%\%Nazwa_projektu2%
If !FileExist(Dane2)
{
	FileCreateDir, %Dane2%
	FileCreateDir, %Dane2%/Branza
	FileCreateDir, %Dane2%/Licznik
	FileCreateDir, %Dane2%/Rodzaje_dok
	FileCreateDir, %Dane2%/Uprawnienia
	Gosub, starter_dane
}
If !FileExist(Roboczy2)
	FileCreateDir, %Roboczy2% 
If !FileExist(Wzor2)
{
	FileCreateDir, %Wzor2% 
	Gosub, starter_wzor
}
If !FileExist(Repozytorium2)
	FileCreateDir, %Repozytorium2%  
If !FileExist(Dokumentacja2)
	FileCreateDir, %Dokumentacja2% 

IF (ukryc_DMS = 1)
{
	FileSetAttrib, +H, %Dane2%
	FileSetAttrib, +H, %Wzor2%
	FileSetAttrib, +H, %Repozytorium2%
}

FileAppend, %Dane2%`n%Roboczy2%`n%Wzor2%`n%Repozytorium2%`n%Dokumentacja2%`n%dir_projekty%, %nowy_dodaj_dr%\info_DMS.txt
FileSetAttrib, +H, %nowy_dodaj_dr%\info_DMS.txt
Plik_info_projekt = %Repozytorium2%\Info.txt
If !FileExist(Plik_info_projekt)
	FileAppend, %Info_projekt2%, %Repozytorium2%\Info.txt, UTF-8
FileCopy, %A_ScriptFullPath%, %nowy_dodaj_dr%, 0

SplashTextOff

MsgBox, 64, Informacja, Projekt został utworzony

projekt_dir := "Wybierz projekt||"
GuiControl, przelacz: , projekt_dir, |%projekt_dir% 
Loop, files, %dir_projekty%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,projekt_dir
	projekt_dir_2 = %A_LoopFileFullPath%\info_DMS.txt
	IF FileExist(projekt_dir_2)
		GuiControl, przelacz: , projekt_dir, %projekt_dir% 
}

Gui, dodaj:Destroy
return

utworz2_dr_pr:
Gui, dodaj2:Submit, noHide

SplashTextOn, 180,40, Proces, Trwa proces...

nowy_dodaj_dr = %dir_projekty%\%Nazwa_projektu%
If !FileExist(nowy_dodaj_dr)
	FileCreateDir, %dir_projekty%\%Nazwa_projektu%
If !FileExist(Dane)
	FileCreateDir, %Dane%
If !FileExist(Roboczy)
	FileCreateDir, %Roboczy% 
If (Roboczy != Roboczy_pomocniczy)
	FileCopyDir, %Roboczy_pomocniczy%, %Roboczy%, 1
If !FileExist(Wzor)
	FileCreateDir, %Wzor% 
If !FileExist(Repozytorium)
	FileCreateDir, %Repozytorium%  	
If (Repozytorium != Repozytorium_pomocniczy)
	FileCopyDir, %Repozytorium_pomocniczy%, %Repozytorium%, 1
If !FileExist(Dokumentacja)
	FileCreateDir, %Dokumentacja% 

IF (ukryc_DMS = 1)
{
	FileSetAttrib, +H, %Dane%
	FileSetAttrib, +H, %Wzor%
	FileSetAttrib, +H, %Repozytorium%
}
IF (ukryc_DMS = 0)
{
	FileSetAttrib, -H, %Dane%
	FileSetAttrib, -H, %Wzor%
	FileSetAttrib, -H, %Repozytorium%
}

FileDelete, %dir_projekty%\%Nazwa_projektu%\info_DMS.txt
FileAppend, %Dane%`n%Roboczy%`n%Wzor%`n%Repozytorium%`n%Dokumentacja%`n%dir_projekty%, %dir_projekty%\%Nazwa_projektu%\info_DMS.txt
FileSetAttrib, +H, %nowy_dodaj_dr%\info_DMS.txt

SplashTextOff

Numer_dok_2 :="Wybierz dokumentację||"
GuiControl, dok: , Numer_dok_2, |%Numer_dok_2% 
Loop, files, %Repozytorium%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	GuiControl, dok: , Numer_dok_2, %Numer_dok_2% 
} 
Numer_dok_2 := "Wybierz dokumentację"
dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%
dok_2 := "Wybierz dokument"
Gosub, radio_filtr_clean
Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

Gui, dodaj2:Destroy
return

zamknij_dr_pr:
Gui, przelacz:Destroy
return

Raport_pro:
Gui, przelacz:Destroy

FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
FormatTime, data_2, A_Now, yyyy_MM_dd_HHmmss

Raport_projektu = Roboczy`n
Loop, files, %Repozytorium%\*, D
{ 
	Status_pro_2 := ""
	Status1_pro := ""
	Status2_pro := ""
	Status3_pro := ""
	
	nr_dok_2 = %A_LoopFileFullPath%
	Loop, files, %nr_dok_2%\*, D
	{ 
		FileReadLine, Spr_stat2, %A_LoopFileFullPath%\info.txt, 6
		If (Spr_stat2 = "Wydany")
			Status1_pro := "Wydana"
		If else (Spr_stat2 = "Akceptacja")
			Status2_pro := "Akceptacja"
		If else (Spr_stat2 != "Wydany" and Spr_stat2 != "Akceptacja")
			Status3_pro := "Roboczy"
	} 
	If (Status1_pro != "")
		Status_pro_2 = %Status1_pro%
	If else (Status2_pro != "")
		Status_pro_2 = %Status2_pro%
	If else (Status3_pro != "")
		Status_pro_2 = %Status3_pro%
	
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	If (Status_pro_2 = "Roboczy" or Status_pro_2 = "Wydana")
		Raport_projektu = %Raport_projektu%`n%A_Tab%%A_Tab%%Status_pro_2%%A_Tab%%A_Tab%%Numer_dok_2%`n`n
	else
		Raport_projektu = %Raport_projektu%`n%A_Tab%%A_Tab%%Status_pro_2%%A_Tab%%Numer_dok_2%`n`n
	
	Loop, files, %nr_dok_2%\*, D
	{
		SplitPath, A_LoopFileFullPath,,,,Numer_dok_3
		FileReadLine, Rewizja_1, %A_LoopFileFullPath%\info.txt, 3
		FileReadLine, Rewizja_2, %A_LoopFileFullPath%\info.txt, 4
		FileReadLine, Status_dok_2 ,%A_LoopFileFullPath%\info.txt, 6
		Rewizja_3 := Format("{01:02}", Rewizja_2)
		If (Status_dok_2 = "Roboczy" or Status_dok_2 = "Wydany")
			Raport_projektu = %Raport_projektu%%A_Tab%%Rewizja_1%.%Rewizja_3%%A_Tab%%Status_dok_2%%A_Tab%%A_Tab%%Numer_dok_3%`n
		else
			Raport_projektu = %Raport_projektu%%A_Tab%%Rewizja_1%.%Rewizja_3%%A_Tab%%Status_dok_2%%A_Tab%%Numer_dok_3%`n
	}
	
} 

Raport_projektu_2 = Dokumentacja\Edytowalne`n
Loop, files, %Dokumentacja%\Edytowalne\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	Raport_projektu_2 = %Raport_projektu_2%`n%A_Tab%%Numer_dok_2%`n`n
	nr_dok_2 = %A_LoopFileFullPath%
	Loop, files, %nr_dok_2%\*, R
	{
		SplitPath, A_LoopFileFullPath,,,dok_3_ext,Numer_dok_3
		If (dok_3_ext != "rar")
			If (dok_3_ext != "zip")
				If (dok_3_ext != "7z")
					Raport_projektu_2 = %Raport_projektu_2%%A_Tab%%Numer_dok_3%`n
	}
}

Raport_projektu_3 = Dokumentacja\Nieedytowalne`n
Loop, files, %Dokumentacja%\Nieedytowalne\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Numer_dok_2
	Raport_projektu_3 = %Raport_projektu_3%`n%A_Tab%%Numer_dok_2%`n`n
	nr_dok_2 = %A_LoopFileFullPath%
	Loop, files, %nr_dok_2%\*, R
	{
		SplitPath, A_LoopFileFullPath,,,,Numer_dok_3
		Raport_projektu_3 = %Raport_projektu_3%%A_Tab%%Numer_dok_3%`n
	}
}

Gui, raport:Add, Text, y10 x10 h20 w400, %data%
Gui, raport:Font, Bold
Gui, raport:Add, Text, y+5 w400 h20, Roboczy:
Gui, raport:Font, 
Gui, raport:Add, Edit, y+10 w400 h120 ReadOnly, %Raport_projektu%
Gui, raport:Font, Bold
Gui, raport:Add, Text, y+10 w340 h20, Dokumentacja\Edytowalne:
Gui, raport:Font, 
Gui, raport:Add, Edit, y+10 w400 h120 ReadOnly, %Raport_projektu_2%
Gui, raport:Font, Bold
Gui, raport:Add, Text, y+10 h20 w340, Dokumentacja\Nieedytowalne:
Gui, raport:Font, 
Gui, raport:Add, Edit, y+10 w400 h120 ReadOnly, %Raport_projektu_3%
Gui, raport:Add, Button, y+10 w60 gexport_raport, Exportuj
Gui, raport:Add, Button, x+10 w60 Default gzamknij_raport, Zamknij

Gui, raport:Show,, Raport
return

export_raport:
Gui, raport:Submit, noHide
FileSelectFolder, raport_dir
If ErrorLevel
	return
FileAppend, %data%`n`n`n%Raport_projektu%`n`n%Raport_projektu_2%`n`n%Raport_projektu_3%`n`n, %raport_dir%\%data_2%_Raport.txt
MsgBox, 64, Informacja, Raport został wyeksportowany
return

zamknij_raport:
Gui, raport:Destroy
return

;_______________________________________________________Dokumenty_____________________________________________

info_dok_2:
Gui, dok:Submit, noHide

if (A_GuiControlEvent = "DoubleClick")
	If (dok_2 != "Wybierz dokument")
	Gosub, otworz_dok_2

Org_dok_2 := ""
Nr_dok_2 := ""
Rewizja_1 := ""
Rewizja_2 := ""
Rewizja_4 := ""
Status_dok := ""
Opis_dok := ""
FileReadLine, Org_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
FileReadLine, Nr_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 2
FileReadLine, Rewizja_1 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 3
FileReadLine, Rewizja_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 4
FileReadLine, Rewizja_4 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 5
FileReadLine, Status_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 6
FileReadLine, dok_ext ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 7
FileReadLine, Opis_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 8
FileRead, His_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
Rewizja_3 := Format("{01:02}", Rewizja_2)		
Rewizja_5 := Format("{01:03}", Rewizja_4)
Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%
If (Rewizja_dok = ".00.000")
	Rewizja_dok := ""

dok_2_spr = %Roboczy%\%Numer_dok_2%\%Org_dok_2%
If !FileExist(dok_2_spr)
	MsgBox, 16, Ostrzeżenie, Niezgodność w pliku: %dok_2% 

GuiControl, dok: , Nr_dok_2, %Nr_dok_2%
GuiControl, dok: , Rewizja_dok, %Rewizja_dok%
GuiControl, dok: , Status_dok, %Status_dok% 
GuiControl, dok: , Opis_dok, %Opis_dok% 
GuiControl, dok: , His_dok, %His_dok%

arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2%

Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,arch_dok_2
	If (arch_dok_2 != "his")
		GuiControl, dok: , arch_dok_2, %arch_dok_2%
}

etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2%

Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.rar, R
{ 
	SplitPath, A_LoopFileFullPath,etransmit_2
	GuiControl, dok: , etransmit_2, %etransmit_2%
} 
Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.zip, R
{ 
	SplitPath, A_LoopFileFullPath,etransmit_2
	GuiControl, dok: , etransmit_2, %etransmit_2%
} 
Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.7z, R
{ 
	SplitPath, A_LoopFileFullPath,etransmit_2
	GuiControl, dok: , etransmit_2, %etransmit_2%
} 

dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 

Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.pdf, R
{ 
	SplitPath, A_LoopFileFullPath,dok_pdf_2
	GuiControl, dok: , dok_pdf_2, %dok_pdf_2% 		
}
dok_pdf_2 := "Wybierz plik PDF"
return

edytuj_dok_2:
Gui, edytuj_dok:Destroy

If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	return
}

Gui, edytuj_dok:Font, Bold
Gui, edytuj_dok:Add,Text, x10 y10 h20 w200 Section, Nazwa dokumentu:
Gui, edytuj_dok:Font,
Gui, edytuj_dok:Add, Edit, y+5 w200 ved_dok_2, %dok_2%
Gui, edytuj_dok:Add, Checkbox, y+10 w200 vchecked_edytuj_klucze, Zmień wewnątrz dokumentu „docx”
Gui, edytuj_dok:Add, Button, y+10 w60 Default ged_dok_2, Zmień

Gui, edytuj_dok:Show,, Edytuj dokument
return

ed_dok_2:
Gui, edytuj_dok:Submit,Nohide

FileReadLine, Org_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
FileReadLine, Nr_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 2
FileReadLine, Rewizja_1 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 3
FileReadLine, Rewizja_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 4
FileReadLine, Rewizja_4 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 5
FileReadLine, Status_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 6
FileReadLine, dok_ext ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 7
Rewizja_3 := Format("{01:02}", Rewizja_2)	

NEW_NUMER_DOKUMENTU = %ed_dok_2%
OLD_NUMER_DOKUMENTU = %Nr_dok_2%

dok_2_spr = %Roboczy%\%Numer_dok_2%\%ed_dok_2%.%dok_ext%
If FileExist(dok_2_spr)
{
	MsgBox, 16, Ostrzeżenie!, Plik o podanej nazwie już istnieje
	return
}
dok_2_spr = %Repozytorium%\%Numer_dok_2%\%ed_dok_2%
If FileExist(dok_2_spr)
{
	MsgBox, 16, Ostrzeżenie!, Plik o podanej nazwie już istnieje
	return
}

FileMove, %Roboczy%\%Numer_dok_2%\%Org_dok_2%, %Roboczy%\%Numer_dok_2%\%ed_dok_2%.%dok_ext%
FileMoveDir, %Repozytorium%\%Numer_dok_2%\%dok_2%, %Repozytorium%\%Numer_dok_2%\%ed_dok_2%

FileDelete, %Repozytorium%\%Numer_dok_2%\%ed_dok_2%\info.txt
FileAppend, %ed_dok_2%.%dok_ext%`n%ed_dok_2%`n%Rewizja_1%`n%Rewizja_3%`n%Rewizja_5%`n%Status_dok%`n%dok_ext%`n, %Repozytorium%\%Numer_dok_2%\%ed_dok_2%\info.txt

FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
FileDelete, %Repozytorium%\%Numer_dok_2%\%ed_dok_2%\his.txt
FileAppend, %Rewizja_1%.%Rewizja_3%.%Rewizja_5%`n%data%`nZmieniono nazwę pliku z %dok_2% na %ed_dok_2% przez %upr_imie%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%ed_dok_2%\his.txt
FileRead, His_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
GuiControl, dok: , His_dok, %His_dok% 	

dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%

Loop, files, %Repozytorium%\%Numer_dok_2%\*%filtr2%*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,dok_2
	FileReadLine, Spr_stat, %A_LoopFileFullPath%\info.txt, 6
	If (filtr3 = "")
		GuiControl, dok: , dok_2, %dok_2%
	If else (Filtr3 = Spr_stat)
		GuiControl, dok: , dok_2, %dok_2%
} 
dok_2 := "Wybierz dokument"


If (dok_ext = "docx" and checked_edytuj_klucze = 1)
{
	SplashTextOn, 180,40, Proces, Trwa proces...
	Path = %Roboczy%\%Numer_dok_2%\%ed_dok_2%.%dok_ext%
	oWord := ComObjCreate("Word.Application")
	oWord.Visible := false
	oWord.Documents.Open(Path)
	FindAndReplace( oWord, OLD_NUMER_DOKUMENTU, NEW_NUMER_DOKUMENTU )
	oWord.Documents.Save()
	oWord.quit()	
	SplashTextOff
}


Gui,edytuj_dok:Destroy
return

pdf_otworz_dok:
Gui, dok:Submit, noHide

if (A_GuiControlEvent = "DoubleClick")
	If (dok_pdf_2 != "Wybierz plik PDF")
		Gosub, otworz_pdf

If (arch_dok_2 != "")
{
	arch_dok_2 := ""
	GuiControl, dok: , arch_dok_2, |%arch_dok_2%
	
	Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*, D
	{ 
		SplitPath, A_LoopFileFullPath,,,,arch_dok_2
		If (arch_dok_2 != "his")
			GuiControl, dok: , arch_dok_2, %arch_dok_2%
	}	
	arch_dok_2 := ""
}

return

arch_otworz_dok:
Gui, dok:Submit, noHide

if (A_GuiControlEvent = "DoubleClick")
	If (arch_dok_2 != "")
		Gosub, Otworz_arch

If (dok_pdf_2 != "Wybierz plik PDF")
{
	dok_pdf_2 := "Wybierz plik PDF||"
	GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
	
	Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.pdf, R
	{ 
		SplitPath, A_LoopFileFullPath,dok_pdf_2
		GuiControl, dok: , dok_pdf_2, %dok_pdf_2% 		
	}
	dok_pdf_2 := "Wybierz plik PDF"
}

return


Archiwizuj_dok:
Gui, dok:Submit, noHide
Gosub, Archiwizuj
FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileAppend, %Rewizja_1%.%Rewizja_3%.%Rewizja_5%`n%data%`nUtworzono archiwum przez %upr_imie%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileRead, His_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
GuiControl, dok: , His_dok, %His_dok% 	
return

Archiwizuj: 
If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	return
}

FormatTime, data_2, A_Now, yyyy_MM_dd_HHmmss

SplashTextOn, 180,40, Proces, Trwa proces...
FileReadLine, Org_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
FileCreateDir, %Repozytorium%\%Numer_dok_2%\%dok_2%\%data_2%_R%Rewizja_1%_%Rewizja_3%
FileCopy, %Roboczy%\%Numer_dok_2%\%Org_dok_2%, %Repozytorium%\%Numer_dok_2%\%dok_2%\%data_2%_R%Rewizja_1%_%Rewizja_3%\%Org_dok_2%
FileSetAttrib, +R, %Repozytorium%\%Numer_dok_2%\%dok_2%\%data_2%_R%Rewizja_1%_%Rewizja_3%\%Org_dok_2%
SplashTextOff

Rewizja_4++
Rewizja_5 := Format("{01:03}", Rewizja_4)

FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\Info.txt
FileAppend, %Org_dok_2%`n%Nr_dok_2%`n%Rewizja_1%`n%Rewizja_3%`n%Rewizja_5%`n%Status_dok%`n%dok_ext%`n, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt

FileReadLine, Org_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
FileReadLine, Nr_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 2
FileReadLine, Rewizja_1 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 3
FileReadLine, Rewizja_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 4
FileReadLine, Rewizja_4 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 5
FileReadLine, Status_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 6
FileReadLine, dok_ext ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 7
FileReadLine, Opis_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 8

Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%

GuiControl, dok: , Nr_dok_2, %Nr_dok_2% 
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
GuiControl, dok: , Status_dok, %Status_dok% 	
GuiControl, dok: , Opis_dok, %Opis_dok% 	

arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 

Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,arch_dok_2
	If (arch_dok_2 != "his")
		GuiControl, dok: , arch_dok_2, %arch_dok_2% 		
} 
return

otworz_dok_2:
Gui, Dok:Submit, noHide

If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	return
}

FileReadLine, dok_2_org ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1

dok_2_spr = %Roboczy%\%Numer_dok_2%\%Org_dok_2%
If !FileExist(dok_2_spr)
{
	MsgBox, 16, Ostrzeżenie, Niezgodność w pliku: %dok_2% 
	return
}

Run, %Roboczy%\%Numer_dok_2%\%dok_2_org%, 1
return

odczyt_dok_2:
Gui, Dok:Submit, noHide

If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	return
}

FileReadLine, dok_2_org ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1

dok_2_spr = %Roboczy%\%Numer_dok_2%\%Org_dok_2%
If !FileExist(dok_2_spr)
{
	MsgBox, 16, Ostrzeżenie, Niezgodność w pliku: %dok_2% 
	return
}

FileSetAttrib, +R, %Roboczy%\%Numer_dok_2%\%dok_2_org%
Run, %Roboczy%\%Numer_dok_2%\%dok_2_org%, 1
Sleep, 1000
FileSetAttrib, -R, %Roboczy%\%Numer_dok_2%\%dok_2_org%
return

usun_dok_2:
Gui, Dok:Submit, noHide
Gui, usun_dok:Destroy

If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	return
}

Gui, usun_dok:Add, Text, x10 y10 h20 w270, Czy na pewno chcesz usunąć zaznaczony dokument?
Gui, usun_dok:Add, Button, y+10 x70 w60 gUsun_dok_3, Usuń
Gui, usun_dok:Add, Button, x+10 w60 Default gUsun_dok_4 , Anuluj

Gui, usun_dok:Show,, Ostrzeżenie!
return

Usun_dok_3:
FileReadLine, dok_org_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
SplashTextOn, 180,40, Proces, Trwa proces...
FileDelete, %Roboczy%\%Numer_dok_2%\%dok_org_2%
FileRemoveDir, %Repozytorium%\%Numer_dok_2%\%dok_2%, 1
SplashTextOff
dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%

Loop, files, %Repozytorium%\%Numer_dok_2%\*%filtr2%*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,dok_2
	FileReadLine, Spr_stat, %A_LoopFileFullPath%\info.txt, 6
	If (filtr3 = "")
		GuiControl, dok: , dok_2, %dok_2%
	If else (Filtr3 = Spr_stat)
		GuiControl, dok: , dok_2, %dok_2%
} 
dok_2 := "Wybierz dokument"

Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

Gui, usun_dok:Destroy
return

Usun_dok_4:
Gui, usun_dok:Destroy
return

wymien_dok_2:
Gui, Dok:Submit, noHide

If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	return
}
If (Status_dok = "Akceptacja")
{
	MsgBox, 16, Ostrzeżenie, Nie można wymienić dokumentu, gdy jego status jest Akceptacja
	return
}
FileReadLine, dok_org_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
ext_2 = %Roboczy%\%Numer_dok_2%\%dok_org_2%
SplitPath, ext_2,,,ext_1
FileSelectFile, dok_new_2,,,,*.%ext_1%
If !ErrorLevel
{
	FileDelete, %Roboczy%\%Numer_dok_2%\%dok_org_2%
	Gosub, Archiwizuj
	SplashTextOn, 180,40, Proces, Trwa proces...
	FileCopy, %dok_new_2%, %Roboczy%\%Numer_dok_2%\%dok_org_2%, 1
	SplashTextOff
	FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
	FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
	FileAppend, %Rewizja_1%.%Rewizja_3%.%Rewizja_5%`n%data%`nPlik został wymieniony przez %upr_imie%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
	FileRead, His_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
	GuiControl, dok: , His_dok, %His_dok% 		
}
return

Nowy_dok:
Gui, stw:Destroy
Gui, Dok:Submit, noHide

If (Numer_dok_2 = "Wybierz dokumentację")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokumentację
	return
}

Lista_branz_dr = %Dane%\Branza
If !FileExist(Lista_branz_dr)
	FileCreateDir, %Dane%\Branza

Numer_dokumetu := ""
branzadok := ""
rodzajdok := ""
Opis_dok := ""

Licznik_dir = %Dnae%\Licznik
IF !FileExist(Licznik_dir)
{
	FileCreateDir, %Licznik_dir%
	FileAppend, 00`n001`n-, %Licznik_dir%\info.txt
}
FileReadLine, Licznik1, %Licznik_dir%\info.txt, 1
FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2
FileReadLine, Licznik3, %Licznik_dir%\info.txt, 3

Num = %licznik1%%licznik2%

Gui, stw:Font, Bold
Gui, stw:Add, Text, x10 y10 w130 h20 Section, Branża:
Gui, stw:Font,
Gui, stw:Add, Button, x+10 w60 vButtonToHide1 gbranza_dodaj, Dodaj
Gui, stw:Add, Button, x+10 w60 vButtonToHide2 gbranza_edytuj gbranza_edytuj, Edytuj
Gui, stw:Add, Button, x+10 w60 vButtonToHide3 gusun_branz, Usuń
Gui, stw:Add, DropDownList, y+5 xs w340 vLista_branz gbranza, %Lista_branz%
Gui, stw:Font, Bold
Gui, stw:Add, Text, y+10 w130 h20, Rodzaj dokumentu:
Gui, stw:Font,
Gui, stw:Add, Button, x+10 w60 vButtonToHide4 gRodz_dok_dodaj, Dodaj
Gui, stw:Add, Button, x+10 w60 vButtonToHide5 gRodz_dok_edytuj, Edytuj
Gui, stw:Add, Button, x+10 w60 vButtonToHide6 gusun_rodz_dok, Usuń
Gui, stw:Add, DropDownList, y+5 xs w340 vLista_dok grodzaj_dokumentu, %Lista_dok%
Gui, stw:Font, Bold
Gui, stw:Add, Text, y+10 w130 h20, Szablon:
Gui, stw:Font,
Gui, stw:Add, Button, x+10 w60 vButtonToHide7 gdodaj_szablon, Dodaj
Gui, stw:Add, Button, x+10 w60 vButtonToHide8 gwymien_szablon, Wymień
Gui, stw:Add, Button, x+10 w60 vButtonToHide9 gusun_szablon, Usuń
Gui, stw:Add, ListBox, y+5 xs w340 h120 vSzablon, %Szablon%
Gui, stw:Font, Bold
Gui, stw:Add, Text, y+10 w200 h20, Numer dokumentu:
Gui, stw:Font,
Gui, stw:Add, Button, x+10 w60 vButtonToMove1 gSpec, Spec
Gui, stw:Add, Button, x+10 w60 vButtonToHide10 gEdytuj_licznik, Edytuj
Gui, stw:Add, Edit, y+5 xs w340 Center vNum gNum, %Num%
Gui, stw:Font, Bold
Gui, stw:Add, Text, y+10 w200 h20, Opis dokumentu:
Gui, stw:Font,
Gui, stw:Add, Edit, y+5 xs w340 Center vOpis_dok, %Opis_dok%
Gui, stw:Font, Bold
Gui, stw:Add, Text, y+10 h20, Oznaczenie branzy:
Gui, stw:Font,
Gui, stw:Add, Edit, y+5 w340 h20 ReadOnly vbranzadok, %branzadok%
Gui, stw:Font, Bold
Gui, stw:Add, Text, y+10 h20, Oznaczenie rodzaju dokumentu:
Gui, stw:Font,
Gui, stw:Add, Edit, y+5 w340 h20 ReadOnly vrodzajdok, %rodzajdok%
Gui, stw:Font, Bold
Gui, stw:Add, Text, y+10 h20, Numer dokumentu:
Gui, stw:Font,
Gui, stw:Add, Edit, y+5 w340 h20 ReadOnly vNumer_dokumetu, %Numer_dokumetu%
Gui, stw:Add, Checkbox, y+10 w340 h20 ReadOnly vchecked_stw_kluczowe, Uzupełnij słowa kluczowe w dokumencie „docx”

Gui, stw:Add, Button, y+10 Default w60 gStworz, Stwórz
Gui, stw:Add, Button, x+10 w60 gZamknij, Zamknij	

If else (upr_uzytkownika = "Rozszerzone" or upr_uzytkownika = "Ograniczone")
{
	GuiControl,stw:Hide,ButtonToHide1
	GuiControl,stw:Hide,ButtonToHide2
	GuiControl,stw:Hide,ButtonToHide3
	GuiControl,stw:Hide,ButtonToHide4
	GuiControl,stw:Hide,ButtonToHide5
	GuiControl,stw:Hide,ButtonToHide6
	GuiControl,stw:Hide,ButtonToHide7
	GuiControl,stw:Hide,ButtonToHide8
	GuiControl,stw:Hide,ButtonToHide9
	GuiControl,stw:Hide,ButtonToHide10
	GuiControl,stw:Hide,ButtonToHide11
	GuiControl,stw:Move,ButtonToMove1, x290
}

Lista_branz := "Wybierz branżę||"
GuiControl, stw: , Lista_branz, |%Lista_branz% 
Loop, files, %Dane%\Branza\*, R
{
	SplitPath, A_LoopFileFullPath,,,,Lista_branz
	GuiControl, stw: , Lista_branz, %Lista_branz% 	
}
Lista_branz := "Wybierz branżę"

Lista_dok := "Wybierz rodzaj dokumentu||"
GuiControl, stw: , Lista_dok, |%Lista_dok% 
Loop, files, %Wzor%\*, D
{
	SplitPath, A_LoopFileFullPath,,,,Lista_dok
	GuiControl, stw: , Lista_dok, %Lista_dok% 
} 
Lista_dok := "Wybierz rodzaj dokumentu"

Szablon := "Wybierz szablon||"
GuiControl, stw: , Szablon, |%Szablon% 

Gui, stw:Show,,Stwórz nowy dokument
return

Edytuj_licznik:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide
Gui, Num1:Destroy

Gui, Num1:Add, Text, x10 y10 w200 Section, Pierwszy człon licznika (Spec):
Gui, Num1:Add, Edit, x+10 w130 vlicznik1, %Licznik1%
Gui, Num1:Add, Text, y+10 xs w200, Drugi człon licznika:
Gui, Num1:Add, Edit, x+10 w130 vlicznik2, %Licznik2%
Gui, Num1:Add, Text, y+10 xs w200 Section, Znak łącznika w numerze dokumentu:
Gui, Num1:Add, Edit, x+10 w130 vlicznik3, %Licznik3%
Gui, Num1:Add, Button, y+10 xs w60 gZmien_licznik Default, Zmień

Gui, Num1:Show,, Sposób numeracji dokumentu
return

Zmien_licznik:
Gui, Num1:Submit, noHide

FileDelete, %Licznik_dir%\info.txt
FileAppend, %Licznik1%`n%Licznik2%`n%Licznik3%, %Licznik_dir%\info.txt

FileReadLine, Licznik1, %Licznik_dir%\info.txt, 1
FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2
FileReadLine, Licznik3, %Licznik_dir%\info.txt, 3

Num = %Licznik1%%Licznik2%
GuiControl, stw: , Num, %Num%
GuiControl, imp: , Num, %Num%

Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%

Gui, Num1:Destroy
return

Spec:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide
Gui, Num2:Destroy

If (Lista_branz = "Wybierz branżę")
{
	MsgBox, 16, Ostrzeżenie!, Wybierz branżę
	return
}

Gui, Num2:Font, Bold
Gui, Num2:Add, Text, x10 y10 h20 w270 Section, Branża:
Gui, Num2:Font, 
Gui, Num2:Add, Edit, y+5 w270 h20 ReadOnly Center vLista_branz2, %Lista_branz%
Gui, Num2:Font, Bold
Gui, Num2:Add, Text, y+10 xs w270 Section, Oznaczenie specjalizacji:
Gui, Num2:Font, 
Gui, Num2:Add, Edit, y+10 w270 Center h20 ReadOnly vlicznik1, %licznik1% 
Gui, Num2:Font, Bold
Gui, Num2:Add,Text, y+10 h20 w270, Dostępne specjalizacje
Gui, Num2:Font, 
Gui, Num2:Add, ListBox, y+5 w270 h200 vspec_list glicznik1, %spec_list%
Gui, Num2:Add, Button, y+10 xs w60 gustaw_licznik1 Default, Ustaw
Gui, Num2:Add, Button, x+10 w60 gdodaj_spec, Dodaj
Gui, Num2:Add, Button, x+10 w60 gedytuj_spec, Edytuj
Gui, Num2:Add, Button, x+10 w60 gusun_spec, Usuń

Lista_branz2 = %Lista_branz%
spec_list := "00 - brak specjalizacji||"
GuiControl, Num2: , spec_list, |%spec_list%
Loop, files, %Dane%\licznik\*, R
{ 
	SplitPath, A_LoopFileFullPath,,,,spec_list
	If (spec_list != "info")
	{
		FileReadLine, spec_bran, %A_LoopFileFullPath%,2
		If (Lista_branz2 = spec_bran)	
			GuiControl, Num2: , spec_list, %spec_list% 
	}
} 
spec_list := "00 - brak specjalizacji"
licznik1 := "00"
GuiControl, Num2: , licznik1, %licznik1% 
Gui, Num2:Show,, Specjalizacja dokumentu
return

licznik1:
Gui, Num2:Submit,Nohide
If (spec_list = "00 - brak specjalizacji")
	Licznik1 := "00"
else
{
	spec_dir = %Dane%\licznik\%spec_list%.txt
	FileReadLine, Licznik1, %spec_dir%,1	
}
GuiControl, Num2:, licznik1, %licznik1%
return

usun_spec:
Gui, Num2:Submit, noHide
Gui, Num5:Destroy

If (spec_list = "00 - brak specjalizacji")
{
	MsgBox, 16, Ostrzeżenie!, Nie można usunąć specjalizacji: 00 - brak specjalizacji
	return
}

Gui, Num5:Add, Text, x10 y10 w290 h20, Czy na pewno chcesz usunąć zaznaczony szablon?
Gui, Num5:Add, Button, y+10 xp+80 w60 gusun2_spec, Usuń
Gui, Num5:Add, Button, x+10 w60 Default gusun3_spec , Anuluj

Gui, Num5:Show,, Ostrzeżenie!
return

usun2_spec:
Gui, Num5:Submit, noHide

spec_dir = %Dane%\licznik\%spec_list%.txt
FileDelete, %spec_dir%

spec_list := "00 - brak specjalizacji||"
GuiControl, Num2: , spec_list, |%spec_list%
Loop, files, %Dane%\licznik\*, R
{ 
	SplitPath, A_LoopFileFullPath,,,,spec_list
	If (spec_list != "info")
	{
		FileReadLine, spec_bran, %A_LoopFileFullPath%,2
		If (Lista_branz2 = spec_bran)	
			GuiControl, Num2: , spec_list, %spec_list% 				
		
	}
} 
spec_list := "00 - brak specjalizacji"
licznik1 := "00"
GuiControl, Num2: , licznik1, %licznik1% 

Gui, Num5:Destroy
return

usun3_spec:
Gui, Num5:Destroy
return

edytuj_spec:
Gui, Num2:Submit,Nohide
Gui, Num4:Destroy

If (spec_list = "00 - brak specjalizacji")
{
	MsgBox, 16, Ostrzeżenie!, Nie można edytować specjalizacji: 00 - brak specjalizacji
	return
}

Gui, Num4:Add, Text, x10 y10 h20 w270, Nazwa na liście specjalizacji dokumentu:
Gui, Num4:Add, Edit, y+5 w270 vnum3_naz, %spec_list%
Gui, Num4:Add, Text, y+10 h20 w270, Oznaczenie specjalizacji dokumentu:
Gui, Num4:Add, Edit, y+5 w270 vnum3_ozn, %licznik1% 
Gui, Num4:Add, Button, y+10 w60 Default gdodaj2_spec, Zmień

Gui, Num4:Show,,Edytuj specjalizację
return

ustaw_licznik1:
Gui, Num2:Submit,Nohide

FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2

Num = %licznik1%%licznik2%
Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
Nowa_dok_info=%Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
Loop,
	If FileExist(Nowa_dok_info)
	{
			Num_count := StrLen(Num)
			Num_count := Format("{01:02}", Num_count)
			Num += 1
			Num_count = {01:%Num_count%}
			Num := Format(Num_count, Num)
			Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
			Nowa_dok_info = %Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
			Nowa_dok = %Roboczy%\%Numer_dok_2%\%Numer_dokumetu%.%dok_import_ext%
}
else
	Break			
GuiControl, stw: , Num, %Num%
GuiControl, imp: , Num, %Num%
Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%
Gui, Num2:Destroy
return

dodaj_spec:
Gui, Num2:Submit,Nohide
Gui, Num3:Destroy

Gui, Num3:Add, Text, x10 y10 h20 w270, Nazwa na liście specjalizacji dokumentu:
Gui, Num3:Add, Edit, y+5 w270 vnum3_naz,
Gui, Num3:Add, Text, y+10 h20 w270, Oznaczenie specjalizacji dokumentu:
Gui, Num3:Add, Edit, y+5 w270 vnum3_ozn,
Gui, Num3:Add, Button, y+10 w60 Default gdodaj2_spec, Dodaj

Gui, Num3:Show,,Dodaj specjalizację
return

dodaj2_spec:
Gui, Num3:Submit,Nohide
Gui, Num4:Submit,Nohide

num3_spr = %Dane%\Licznik\%num3_naz%.txt
FileDelete, %num3_spr%
FileAppend, %num3_ozn%`n%Lista_branz%, %num3_spr%

spec_list := "00 - brak specjalizacji||"
GuiControl, Num2: , spec_list, |%spec_list%
Loop, files, %Dane%\licznik\*, R
{ 
	SplitPath, A_LoopFileFullPath,,,,spec_list
	If (spec_list != "info")
	{
		FileReadLine, spec_bran, %A_LoopFileFullPath%,2
		If (Lista_branz2 = spec_bran)	
			GuiControl, Num2: , spec_list, %spec_list% 				
		
	}
} 
spec_list := "00 - brak specjalizacji"
licznik1 := "00"
GuiControl, Num2: , licznik1, %licznik1% 

Gui, Num3:Destroy
Gui, Num4:Destroy
return

usun_szablon:
Gui, stw:Submit, noHide
Gui, S1:Destroy

If (Lista_dok = "Wybierz szablon")
{
	MsgBox, 16, Ostrzeżenie, Wybierz szablon
	return
}

Gui, S1:Add, Text, x10 y10 w290 h20, Czy na pewno chcesz usunąć zaznaczony szablon?
Gui, S1:Add, Button, y+10 xp+80 w60 gusun2_szablon, Usuń
Gui, S1:Add, Button, x+10 w60 Default gusun3_szablon , Anuluj

Gui, S1:Show,, Ostrzeżenie!
return

usun2_szablon:
Gui, S1:Submit, noHide

SplashTextOn, 180,40, Proces, Trwa proces...
FileDelete, %Wzor%\%Lista_dok%\%Szablon%
SplashTextOff

Szablon := "Wybierz szablon||"
GuiControl, stw: , Szablon, |%Szablon% 

Loop, files, %Wzor%\%Lista_dok%\*
	If (A_LoopFileName != "Info.txt")
	{ 
		SplitPath, A_LoopFileFullPath,Szablon
		GuiControl, stw: , Szablon, %Szablon% 
	} 

Gui, S1:Destroy
return

usun3_szablon:
Gui, S1:Destroy
return

wymien_szablon:
Gui, Dok:Submit, noHide

If (Szablon = "Wybierz szablon")
{
	MsgBox, 16, Ostrzeżenie, Wybierz szablon
	return
}

FileSelectFile, new_szablon
SplitPath, new_szablon,,,new_szablon_ext
If !ErrorLevel
{
	Szablon3 = %Wzor%\%Lista_dok%\%Szablon%
	SplitPath, Szablon3,,,,Szablon2
	SplashTextOn, 180,40, Proces, Trwa proces...
	FileDelete, %Szablon3%
	FileCopy, %new_szablon%, %Wzor%\%Lista_dok%\%Szablon2%.%new_szablon_ext%, 1
	SplashTextOff
}

Szablon := "Wybierz szablon||"
GuiControl, stw: , Szablon, |%Szablon% 

Loop, files, %Wzor%\%Lista_dok%\*
	If (A_LoopFileName != "Info.txt")
	{ 
		SplitPath, A_LoopFileFullPath,Szablon
		GuiControl, stw: , Szablon, %Szablon% 
	} 
return

dodaj_szablon:
Gui, stw:Submit, noHide

If (Lista_dok = "Wybierz rodzaj dokumentu")
{
	MsgBox, 16, Ostrzeżenie, Wybierz rodzaj dokumentu
	return
}

FileSelectFile, dok_szablon
SplitPath, dok_szablon,dok_szablon2
SplashTextOn, 180,40, Proces, Trwa proces...
FileCopy, %dok_szablon%, %Wzor%\%Lista_dok%\%dok_szablon2%, 1
SplashTextOff

Szablon := "Wybierz szablon||"
GuiControl, stw: , Szablon, |%Szablon% 

Loop, files, %Wzor%\%Lista_dok%\*
	If (A_LoopFileName != "Info.txt")
	{ 
		SplitPath, A_LoopFileFullPath,Szablon
		GuiControl, stw: , Szablon, %Szablon% 
	} 

return

usun_rodz_dok:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide
Gui, R1_rodz:Destroy

If (Lista_dok = "Wybierz rodzaj dokumentu")
{
	MsgBox, 16, Ostrzeżenie, Wybierz rodzaj dokumentu
	return
}

Gui, R1_rodz:Add, Text, x10 y10 h30 w290, Czy na pewno chcesz usunąć rodzaj dokumentacji wraz z przynależnymi szablonami?
Gui, R1_rodz:Add, Button, y+10 xp+80 w60 gusun_rodz2_dok, Usuń
Gui, R1_rodz:Add, Button, x+10 w60 Default gusun_rodz3_dok , Anuluj

Gui, R1_rodz:Show,, Ostrzeżenie!
return

usun_rodz2_dok:
Gui, R1_rodz:Submit, noHide

FileRemoveDir, %Wzor%\%Lista_dok%, 1

Lista_dok := "Wybierz rodzaj dokumentu||"
GuiControl, stw: , Lista_dok, |%Lista_dok% 
GuiControl, imp: , Lista_dok, |%Lista_dok% 
Loop, files, %Wzor%\*, D
{
	SplitPath, A_LoopFileFullPath,,,,Lista_dok
	GuiControl, stw: , Lista_dok, %Lista_dok% 
	GuiControl, imp: , Lista_dok, %Lista_dok% 
} 

Szablon := "Wybierz szablon||"
GuiControl, stw: , Szablon, |%Szablon% 
GuiControl, imp: , Szablon, |%Szablon% 
rodzajdok := ""
GuiControl, stw: , rodzajdok, %rodzajdok%
GuiControl, imp: , rodzajdok, %rodzajdok%
Numer_dokumetu := ""
GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%

Gui, R1_rodz:Destroy
return

usun_rodz3_dok:
Gui, R1_rodz:Destroy
return

Rodz_dok_dodaj:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide
Gui, Rodz_dok:Destroy

Gui, Rodz_dok:Add, Text, x10 y10 h20 w150 Section, Nazwa rodzaju dokumentu:
Gui, Rodz_dok:Add, Edit, x+10 w130 vrodz_dok, 
Gui, Rodz_dok:Add, Text, y+10 h20 xs w150, Skrót rodzaju dokumentu:
Gui, Rodz_dok:Add, Edit, x+10 w130 vskr_rodz_dok,
Gui, Rodz_dok:Add, Button, y+10 xs w60 gdod_rodz_dok Default, Dodaj

Gui, Rodz_dok:Show,, Dodaj rodzaj dokumentu
return

Rodz_dok_edytuj:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide
Gui, Rodz2_dok:Destroy

If (Lista_dok = "Wybierz rodzaj dokumentu")
{
	MsgBox, 16, Ostrzeżenie, Wybierz rodzaj dokumentu
	return
}

Gui, Rodz2_dok:Add, Text, x10 y10 h20 w150 Section, Nazwa rodzaju dokumentu:
Gui, Rodz2_dok:Add, Edit, x+10 w130 vrodz_dok, %Lista_dok%
Gui, Rodz2_dok:Add, Text, y+10 xs h20 w150, Skrót rodzaju dokumentu:
Gui, Rodz2_dok:Add, Edit, x+10 w130 vskr_rodz_dok, %rodzajdok%
Gui, Rodz2_dok:Add, Button, y+10 xs w60 ged_rodz_dok Default, Edytuj

Gui, Rodz2_dok:Show,, Edytuj rodzaj dokumentu
return

ed_rodz_dok:
Gui, Rodz2_dok:Submit, noHide
If (rodz_dok = "")
{
	MsgBox, 16, Ostrzeżenie, Nadaj nazwę rodzaju dokumentu
	return
}

FileRemoveDir, %Wzor%\%Lista_dok%, 1
FileCreateDir, %Wzor%\%rodz_dok%
FileDelete, %Wzor%\%rodz_dok%\info.txt
FileAppend, %skr_rodz_dok%, %Wzor%\%rodz_dok%\info.txt

Lista_dok := "Wybierz rodzaj dokumentu||"
GuiControl, stw: , Lista_dok, |%Lista_dok% 
GuiControl, imp: , Lista_dok, |%Lista_dok% 
Loop, files, %Wzor%\*, D
{
	SplitPath, A_LoopFileFullPath,,,,Lista_dok
	GuiControl, stw: , Lista_dok, %Lista_dok% 
	GuiControl, imp: , Lista_dok, %Lista_dok% 
} 

Szablon := "Wybierz szablon||"
GuiControl, stw: , Szablon, |%Szablon% 
GuiControl, imp: , Szablon, |%Szablon% 
rodzajdok := ""
GuiControl, stw: , rodzajdok, %rodzajdok%
GuiControl, imp: , rodzajdok, %rodzajdok%
Numer_dokumetu := ""
GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%

Gui, Rodz2_dok:Destroy
return

dod_rodz_dok:
Gui, Rodz_dok:Submit, noHide
If (rodz_dok = "")
{
	MsgBox, 16, Ostrzeżenie, Nadaj nazwę rodzaju dokumentu
	return
}

FileCreateDir, %Wzor%\%rodz_dok%
FileDelete, %Wzor%\%rodz_dok%\info.txt
FileAppend, %skr_rodz_dok%, %Wzor%\%rodz_dok%\info.txt

Lista_dok := "Wybierz rodzaj dokumentu||"
GuiControl, stw: , Lista_dok, |%Lista_dok% 
GuiControl, imp: , Lista_dok, |%Lista_dok% 
Loop, files, %Wzor%\*, D
{
	SplitPath, A_LoopFileFullPath,,,,Lista_dok
	GuiControl, stw: , Lista_dok, %Lista_dok% 
	GuiControl, imp: , Lista_dok, %Lista_dok% 
} 

Szablon := "Wybierz szablon||"
GuiControl, stw: , Szablon, |%Szablon% 
GuiControl, imp: , Szablon, |%Szablon% 
rodzajdok := ""
GuiControl, stw: , rodzajdok, %rodzajdok%
GuiControl, imp: , rodzajdok, %rodzajdok%
Numer_dokumetu := ""
GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%

Gui, Rodz_dok:Destroy
return

usun_branz:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide
Gui, R1_br:Destroy

If (Lista_branz = "Wybierz branżę")
{
	MsgBox, 16, Ostrzeżenie, Wybierz branżę
	return
}

Gui, R1_br:Add, Text, x10 y10 w290, Czy na pewno chcesz usunąć branżę?
Gui, R1_br:Add, Button, y+10 xp+80 w60 gusun_branz_2, Usuń
Gui, R1_br:Add, Button, x+10 w60 Default gusun_branz_3 , Anuluj

Gui, R1_br:Show,, Ostrzeżenie!
return

usun_branz_2:
Gui, R1_br:Submit, noHide

FileDelete, %Dane%\Branza\%Lista_branz%.txt

Lista_branz := "Wybierz branżę||"
GuiControl, stw: , Lista_branz, |%Lista_branz% 
GuiControl, imp: , Lista_branz, |%Lista_branz% 
Loop, files, %Dane%\Branza\*, R
{
	SplitPath, A_LoopFileFullPath,,,,Lista_branz
	GuiControl, stw: , Lista_branz, %Lista_branz% 	
	GuiControl, imp: , Lista_branz, %Lista_branz% 	
}

branzadok := ""
GuiControl, stw: , branzadok, %branzadok%
GuiControl, imp: , branzadok, %branzadok%
Numer_dokumetu := ""
GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%

Gui, R1_br:Destroy
return

usun_branz_3:
Gui, R1_br:Destroy
return

branza_dodaj:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide
Gui, dod_branza:Destroy

Gui, dod_branza:Add, Text, x10 y10 w150 Section, Nazwa branży:
Gui, dod_branza:Add, Edit, x+10 w130 vnaz_branza, 
Gui, dod_branza:Add, Text, y+10 xs w150, Skrót branży:
Gui, dod_branza:Add, Edit, x+10 w130 vskr_branza,
Gui, dod_branza:Add, Button, y+10 xs w60 gdod_branza Default, Dodaj

Gui, dod_branza:Show,, Dodaj branżę
return

branza_edytuj:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide
Gui, dod_2_branza:Destroy

If (Lista_branz = "Wybierz branżę")
{
	MsgBox, 16, Ostrzeżenie, Wybierz branżę
	return
}

Gui, dod_2_branza:Add, Text, x10 y10 w150 Section, Nazwa branży:
Gui, dod_2_branza:Add, Edit, x+10 w130 vnaz_branza, %Lista_branz%
Gui, dod_2_branza:Add, Text, y+15 xs w150, Skrót branży:
Gui, dod_2_branza:Add, Edit, x+10 w130 vskr_branza, %branzadok%
Gui, dod_2_branza:Add, Button, y+10 xs w60 ged_branza Default, Edytuj

Gui, dod_2_branza:Show,, Edytuj branżę
return

ed_branza:
Gui, dod_2_branza:Submit, noHide
If (naz_branza = "")
{
	MsgBox, 16, Ostrzeżenie, Nadaj nazwę branży
	return
}

FileDelete, %Dane%\Branza\%Lista_branz%.txt
FileDelete, %Dane%\Branza\%naz_branza%.txt
FileAppend, %skr_branza%, %Dane%\Branza\%naz_branza%.txt

Lista_branz := "Wybierz branżę||"
GuiControl, stw: , Lista_branz, |%Lista_branz% 
GuiControl, imp: , Lista_branz, |%Lista_branz% 
Loop, files, %Dane%\Branza\*, R
{
	SplitPath, A_LoopFileFullPath,,,,Lista_branz
	GuiControl, stw: , Lista_branz, %Lista_branz% 	
	GuiControl, imp: , Lista_branz, %Lista_branz% 	
}

branzadok := ""
GuiControl, stw: , branzadok, %branzadok%
GuiControl, imp: , branzadok, %branzadok%
Numer_dokumetu := ""
GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%

Gui, dod_2_branza:Destroy
return

dod_branza:
Gui, dod_branza:Submit, noHide
If (naz_branza = "")
{
	MsgBox, 16, Ostrzeżenie, Nadaj nazwę branży
	return
}

FileDelete, %Dane%\Branza\%naz_branza%.txt
FileAppend, %skr_branza%, %Dane%\Branza\%naz_branza%.txt

Lista_branz := "Wybierz branżę||"
GuiControl, stw: , Lista_branz, |%Lista_branz% 
GuiControl, imp: , Lista_branz, |%Lista_branz% 
Loop, files, %Dane%\Branza\*, R
{
	SplitPath, A_LoopFileFullPath,,,,Lista_branz
	GuiControl, stw: , Lista_branz, %Lista_branz% 	
	GuiControl, imp: , Lista_branz, %Lista_branz% 	
}

branzadok := ""
GuiControl, stw: , branzadok, %branzadok%
GuiControl, imp: , branzadok, %branzadok%
Numer_dokumetu := ""
GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%

Gui, dod_branza:Destroy
return

import_dok_2:
Gui, Dok:Submit, noHide
Gui, imp:Destroy

If (Numer_dok_2 = "Wybierz dokumentację")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokumentację
	return
}

FileSelectFile, dok_import_path, M

If ErrorLevel
{
	return
}
SplashTextOn, 180,40, Proces, Trwa proces...
Loop, Parse, dok_import_path, `n
{
	
	If (A_Index = 1)
		dok_import2_path = % A_LoopField
	If (A_Index = 2)
	{
		dok_import = % A_LoopField
		FileDelete, %Dane%/Temp/info_%A_UserName%.txt
		FileAppend,%dok_import2_path%`n%dok_import%, %Dane%/Temp/info_%A_UserName%.txt		
	}
	else
	{
		dok_import := "Wiele plików"	
		dok_import2 = % A_LoopField
		FileAppend,`n%dok_import2%, %Dane%/Temp/info_%A_UserName%.txt	
	}
}

Numer_dokumetu := ""
branzadok := ""
rodzajdok := ""
Opis_dok := ""

Licznik_dir = %Dane%\Licznik
IF !FileExist(Licznik_dir)
{
	FileCreateDir, %Licznik_dir%
	FileAppend, 00`n001`n-, %Licznik_dir%\info.txt
}

FileReadLine, Licznik1, %Licznik_dir%\info.txt, 1
FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2
FileReadLine, Licznik3, %Licznik_dir%\info.txt, 3

Num = %licznik1%%licznik2%

Gui, imp:Font, Bold
Gui, imp:Add, Text, x10 y10 h20 Section w130, Branża:
Gui, imp:Font,
Gui, imp:Add, Button, x+10 w60 vButtonToHide1 gbranza_dodaj, Dodaj
Gui, imp:Add, Button, x+10 w60 vButtonToHide2 gbranza_edytuj gbranza_edytuj, Edytuj
Gui, imp:Add, Button, x+10 w60 vButtonToHide3 gusun_branz, Usuń
Gui, imp:Add, DropDownList, y+5 xs w340 vLista_branz gbranza, %Lista_branz%
Gui, imp:Font, Bold
Gui, imp:Add, Text, y+10 w130 h20, Rodzaj dokumentu:
Gui, imp:Font,
Gui, imp:Add, Button, x+10 w60 vButtonToHide4 gRodz_dok_dodaj, Dodaj
Gui, imp:Add, Button, x+10 w60 vButtonToHide5 gRodz_dok_edytuj, Edytuj
Gui, imp:Add, Button, x+10 w60 vButtonToHide6 gusun_rodz_dok, Usuń
Gui, imp:Add, DropDownList, y+5 xs w340 vLista_dok grodzaj_dokumentu, %Lista_dok%
Gui, imp:Font, Bold
Gui, imp:Add, Text, y+10 w270 h20, Importowany plik:
Gui, imp:Font,
Gui, imp:Add, Button, x+10 w60 gImport_zmien, Zmień
Gui, imp:Add, Edit, xs y+5 w340 h20 ReadOnly vSzablon, %dok_import%
Gui, imp:Font, Bold
Gui, imp:Add, Text, y+10 w200, Numer dokumentu:
Gui, imp:Font,
Gui, imp:Add, Button, x+10 w60 h20 vButtonToMove1 gSpec, Spec
Gui, imp:Add, Button, x+10 w60 vButtonToHide7 gEdytuj_licznik, Edytuj
Gui, imp:Add, Edit, y+5 xs w340 Center vNum gNum, %Num%
Gui, imp:Font, Bold
Gui, imp:Add, Text, y+10 w200 h20, Opis dokumentu:
Gui, imp:Font,
Gui, imp:Add, Edit, y+5 xs w340 Center vOpis_dok, %Opis_dok%
Gui, imp:Font, Bold
Gui, imp:Add, Text, y+10 h20, Oznaczenie branzy:
Gui, imp:Font,
Gui, imp:Add, Edit, y+5 w340 h20 ReadOnly vbranzadok, %branzadok%
Gui, imp:Font, Bold
Gui, imp:Add, Text, y+10 h20, Oznaczenie rodzaju dokumentu:
Gui, imp:Font,
Gui, imp:Add, Edit, y+5 w340 h20 ReadOnly vrodzajdok, %rodzajdok%
Gui, imp:Font, Bold
Gui, imp:Add, Text, y+10 h20, Numer dokumentu:
Gui, imp:Font,
Gui, imp:Add, Edit, y+5 w340 h20 ReadOnly vNumer_dokumetu, %Numer_dokumetu%
Gui, imp:Add, Checkbox, y+10 w340 h20 ReadOnly vchecked_imp gOrg_numer, Zachowaj oryginalną nazwę dokumentu
Gui, imp:Add, Checkbox, y+10 w340 h20 ReadOnly vchecked_imp_kluczowe, Uzupełnij słowa kluczowe w dokumencie „docx”

Gui, imp:Add, Button, y+10 Default w60 gStworz_Import, Importuj
Gui, imp:Add, Button, x+10 w60 gZamknij, Zamknij	

If else (upr_uzytkownika = "Rozszerzone" or upr_uzytkownika = "Ograniczone")
{
	GuiControl,imp:Hide,ButtonToHide1
	GuiControl,imp:Hide,ButtonToHide2
	GuiControl,imp:Hide,ButtonToHide3
	GuiControl,imp:Hide,ButtonToHide4
	GuiControl,imp:Hide,ButtonToHide5
	GuiControl,imp:Hide,ButtonToHide6
	GuiControl,imp:Hide,ButtonToHide7
	GuiControl,imp:Move,ButtonToMove1, x290
}

Lista_branz := "Wybierz branżę||"
GuiControl, imp: , Lista_branz, |%Lista_branz% 
Loop, files, %Dane%\Branza\*, R
{ 
	SplitPath, A_LoopFileFullPath,,,,Lista_branz
	GuiControl, imp: , Lista_branz, %Lista_branz% 
}
Lista_branz := "Wybierz branżę"

Lista_dok := "Wybierz rodzaj dokumentacji||"
GuiControl, imp: , Lista_dok, |%Lista_dok% 
Loop, files, %Wzor%\*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,Lista_dok
	GuiControl, imp: , Lista_dok, %Lista_dok% 
}
Lista_dok := "Wybierz rodzaj dokumentacji"

SplashTextOff
Gui, imp:Show,,Zaimportuj nowy dokument
return

Import_zmien:
Gui, Dok:Submit, noHide
FileSelectFile, dok_import_path, M

If ErrorLevel
{
	return
}
Loop, Parse, dok_import_path, `n
{
	
	If (A_Index = 1)
		dok_import2_path = % A_LoopField
	If (A_Index = 2)
	{
		dok_import = % A_LoopField
		FileDelete, %Dane%/Temp/info_%A_UserName%.txt
		FileAppend,%dok_import2_path%`n%dok_import%, %Dane%/Temp/info_%A_UserName%.txt		
	}
	else
	{
		dok_import := "Wiele plików"	
		dok_import2 = % A_LoopField
		FileAppend,`n%dok_import2%, %Dane%/Temp/info_%A_UserName%.txt	
	}
}
If (checked_imp = 0)
{
	Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
	GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%
	GuiControl, imp: , szablon, %dok_import%
	return
}

Loop, Parse, dok_import_path, `n
{
	If (A_Index = 1)
		dok_import2_path = % A_LoopField
	If (A_Index = 2)
	{
		Numer_dokumetu = % A_LoopField	
	}
	else
	{
		Numer_dokumetu := "Wiele nazw"
	}
}
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%
GuiControl, imp: , szablon, %dok_import% 
return

Org_numer:
Gui, imp:Submit, noHide
If (checked_imp = 0)
{
	Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
	GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
	GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%
	return
}
	Loop, Parse, dok_import_path, `n
	{
		
		If (A_Index = 1)
			dok_import2_path = % A_LoopField
		If (A_Index = 2)
		{
			Numer_dokumetu = % A_LoopField	
		}
		else
		{
			Numer_dokumetu := "Wiele nazw"
		}
}
GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%
return

branza:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide

branzadok := ""
FileReadLine, branzadok, %Dane%\Branza\%Lista_branz%.txt, 1
GuiControl, stw: , branzadok, %branzadok%
GuiControl, imp: , branzadok, %branzadok%

If (checked_imp != 1)
{
	
	FileReadLine, Licznik1, %Licznik_dir%\info.txt, 1 
	FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2
	Num = %licznik1%%licznik2%
	Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
	Nowa_dok_info=%Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
	Loop,
		If FileExist(Nowa_dok_info)
		{
			Num_count := StrLen(Num)
			Num_count := Format("{01:02}", Num_count)
			Num += 1
			Num_count = {01:%Num_count%}
			Num := Format(Num_count, Num)
			Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
			Nowa_dok_info = %Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
			Nowa_dok = %Roboczy%\%Numer_dok_2%\%Numer_dokumetu%.%dok_import_ext%
		}
	else
		Break			
	GuiControl, stw: , Num, %Num%
	GuiControl, imp: , Num, %Num%
	GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
	GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%	
}

return

rodzaj_dokumentu:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide

rodzajdok := ""
FileReadLine, rodzajdok, %Wzor%\%Lista_dok%\Info.txt, 1
GuiControl, stw: , rodzajdok, %rodzajdok%
GuiControl, imp: , rodzajdok, %rodzajdok%

If (checked_imp != 1)
{
	
	FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2
	Num = %licznik1%%licznik2%
	Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
	Nowa_dok_info=%Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
	Loop,
		If FileExist(Nowa_dok_info)
		{
			Num_count := StrLen(Num)
			Num_count := Format("{01:02}", Num_count)
			Num += 1
			Num_count = {01:%Num_count%}
			Num := Format(Num_count, Num)
			Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
			Nowa_dok_info = %Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
			Nowa_dok = %Roboczy%\%Numer_dok_2%\%Numer_dokumetu%.%dok_import_ext%
		}
	else
		Break			
	GuiControl, stw: , Num, %Num%
	GuiControl, imp: , Num, %Num%
	GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
	GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%
	
	Szablon := "Wybierz szablon||"
	GuiControl, stw: , Szablon, |%Szablon% 
	
	Loop, files, %Wzor%\%Lista_dok%\*
		If (A_LoopFileName != "Info.txt")
		{ 
			SplitPath, A_LoopFileFullPath,Szablon
			GuiControl, stw: , Szablon, %Szablon% 
		} 
}
return

Zamknij:
Gui, stw:Destroy
Gui, imp:Destroy
return

Stworz:
Gui, stw:Submit, noHide

If (branzadok = "")
{
	MsgBox, Ostrzeżenie!, Wybierz branżę
	return
}
If (rodzajdok = "")
{
	MsgBox, Ostrzeżenie!, Wybierz rodzaj dokumentu
	return
}
If (Szablon = "Wybierz szablon")
{
	MsgBox, Ostrzeżenie!, Wybierz szablon
	return
}

znaki_spec_spr = %Numer_dokumetu%
Gosub, znaki_spec
If (znak_test = 1)
	return

Nowa_dok_info = %Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
Szablon3 = %Wzor%\%Lista_dok%\%Szablon%
SplitPath, Szablon3,,,ext
Nowa_dok = %Roboczy%\%Numer_dok_2%\%Numer_dokumetu%.%ext%
FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss

If FileExist(Nowa_dok_info)
	MsgBox, 16, Ostrzeżenie!, Plik już istnieje
else
{
	SplashTextOn, 180,40, Proces, Trwa proces...
	FileCopy, %Wzor%\%Lista_dok%\%Szablon%, %Nowa_dok%, 1
	
	FileCreateDir, %Nowa_dok_info%
	
	FileDelete, %Nowa_dok_info%\his.txt
	FileAppend, 0.01.001`n%data%`nPlik został utworzony przez: %upr_imie%`n`n, %Nowa_dok_info%\his.txt
	
	FileDelete, %Nowa_dok_info%\info.txt
	FileAppend, %Numer_dokumetu%.%ext%`n%Numer_dokumetu%`n0`n01`n001`nRoboczy`n%ext%`n%Opis_dok%`n%Lista_branz%`n, %Nowa_dok_info%\info.txt
	
	IF (ext = "docx" and checked_stw_kluczowe = 1)
	{
; Dane na podstawie Informacji o projekcie
		
		OLD_NUMER_PROJEKTU := "NEW_NUMER_PROJEKTU"
		FileReadLine, NEW_NUMER_PROJEKTU, %Repozytorium%\info.txt, 1
		SPR_NUMER_PROJEKTU := SubStr(NEW_NUMER_PROJEKTU,1,17) 
		SPR2_NUMER_PROJEKTU := "Numer projektu:"A_Tab A_Tab
		If (SPR_NUMER_PROJEKTU = SPR2_NUMER_PROJEKTU)
		{
			NEW_NUMER_PROJEKTU := SubStr(NEW_NUMER_PROJEKTU, 18) 
			If (NEW_NUMER_PROJEKTU = "")
				NEW_NUMER_PROJEKTU := " "
		}
		else
			OLD_NUMER_PROJEKTU := ""
		
		OLD_NAZWA_INWESTYCJI := "NEW_NAZWA_INWESTYCJI"
		FileReadLine, NEW_NAZWA_INWESTYCJI, %Repozytorium%\info.txt, 2
		SPR_NAZWA_INWESTYCJI := SubStr(NEW_NAZWA_INWESTYCJI,1,19) 
		SPR2_NAZWA_INWESTYCJI := "Nazwa inwestycji:"A_Tab A_Tab
		If (SPR_NAZWA_INWESTYCJI = SPR2_NAZWA_INWESTYCJI)
		{
			NEW_NAZWA_INWESTYCJI := SubStr(NEW_NAZWA_INWESTYCJI, 20) 
			If (NEW_NAZWA_INWESTYCJI = "")
				NEW_NAZWA_INWESTYCJI := " "
		}
		else
			OLD_NAZWA_INWESTYCJI := ""
		
		OLD_ADRES_INWESTYCJI := "NEW_ADRES_INWESTYCJI"
		FileReadLine, NEW_ADRES_INWESTYCJI, %Repozytorium%\info.txt, 3
		SPR_ADRES_INWESTYCJI := SubStr(NEW_ADRES_INWESTYCJI,1, 19) 
		SPR2_ADRES_INWESTYCJI := "Adres inwestycji:"A_Tab A_Tab
		If (SPR_ADRES_INWESTYCJI = SPR2_ADRES_INWESTYCJI)
		{
			NEW_ADRES_INWESTYCJI := SubStr(NEW_ADRES_INWESTYCJI, 20) 
			If (NEW_ADRES_INWESTYCJI = "")
				NEW_ADRES_INWESTYCJI := " "
		}
		else
			OLD_ADRES_INWESTYCJI := ""
		
		OLD_NAZWA_ZAMAWIAJACEGO := "NEW_NAZWA_ZAMAWIAJACEGO"
		FileReadLine, NEW_NAZWA_ZAMAWIAJACEGO, %Repozytorium%\info.txt, 4
		SPR_NAZWA_ZAMAWIAJACEGO := SubStr(NEW_NAZWA_ZAMAWIAJACEGO,1,21) 
		SPR2_NAZWA_ZAMAWIAJACEGO := "Nazwa zamawiającego:"A_Tab
		If (SPR_NAZWA_ZAMAWIAJACEGO = SPR2_NAZWA_ZAMAWIAJACEGO)
		{
			NEW_NAZWA_ZAMAWIAJACEGO := SubStr(NEW_NAZWA_ZAMAWIAJACEGO, 22) 
			If (NEW_NAZWA_ZAMAWIAJACEGO = "")
				NEW_NAZWA_ZAMAWIAJACEGO := " "
		}
		else
			OLD_NAZWA_ZAMAWIAJACEGO := ""
		
		OLD_ADRES_ZAMAWIAJACEGO := "NEW_ADRES_ZAMAWIAJACEGO"
		FileReadLine, NEW_ADRES_ZAMAWIAJACEGO, %Repozytorium%\info.txt, 5
		SPR_ADRES_ZAMAWIAJACEGO := SubStr(NEW_ADRES_ZAMAWIAJACEGO,1,21) 
		SPR2_ADRES_ZAMAWIAJACEGO := "Adres zamawiającego:"A_Tab
		If (SPR_ADRES_ZAMAWIAJACEGO = SPR2_ADRES_ZAMAWIAJACEGO)
		{
			NEW_ADRES_ZAMAWIAJACEGO := SubStr(NEW_ADRES_ZAMAWIAJACEGO, 22) 
			If (NEW_ADRES_ZAMAWIAJACEGO = "")
				NEW_ADRES_ZAMAWIAJACEGO := " "
		}
		else
			OLD_ADRES_ZAMAWIAJACEGO := ""
		
		OLD_NUMER_UMOWY := "NEW_NUMER_UMOWY"
		FileReadLine, NEW_NUMER_UMOWY, %Repozytorium%\info.txt, 6
		SPR_NUMER_UMOWY := SubStr(NEW_NUMER_UMOWY,1,14) 
		SPR2_NUMER_UMOWY := "Numer umowy:"A_Tab A_Tab
		If (SPR_NUMER_UMOWY = SPR2_NUMER_UMOWY)
		{
			NEW_NUMER_UMOWY := SubStr(NEW_NUMER_UMOWY, 15) 
			If (NEW_NUMER_UMOWY = "")
				NEW_NUMER_UMOWY := " "
		}
		else
			OLD_NUMER_UMOWY := ""
		
		OLD_KIEROWNIK_PROJEKTU := "NEW_KIEROWNIK_PROJEKTU"
		FileReadLine, NEW_KIEROWNIK_PROJEKTU, %Repozytorium%\info.txt, 7
		SPR_KIEROWNIK_PROJEKTU := SubStr(NEW_KIEROWNIK_PROJEKTU,1,21) 
		SPR2_KIEROWNIK_PROJEKTU := "Kierownik projektu:"A_Tab A_Tab
		If (SPR_KIEROWNIK_PROJEKTU = SPR2_KIEROWNIK_PROJEKTU)
		{
			NEW_KIEROWNIK_PROJEKTU := SubStr(NEW_KIEROWNIK_PROJEKTU, 22) 
			If (NEW_KIEROWNIK_PROJEKTU = "")
				NEW_KIEROWNIK_PROJEKTU := " "
		}
		else
			OLD_KIEROWNIK_PROJEKTU := ""
		
; Dane na podstawie Informacji o dokumentacji
		
		OLD_RODZAJ_DOKUMENTACJI := "NEW_RODZAJ_DOKUMENTACJI"
		FileReadLine, NEW_RODZAJ_DOKUMENTACJI, %Repozytorium%\%Numer_dok_2%\info.txt, 1
		SPR_RODZAJ_DOKUMENTACJI := SubStr(NEW_RODZAJ_DOKUMENTACJI,1,21) 
		SPR2_RODZAJ_DOKUMENTACJI := "Rodzaj dokumentacji:"A_Tab
		If (SPR_RODZAJ_DOKUMENTACJI = SPR2_RODZAJ_DOKUMENTACJI)
		{
			NEW_RODZAJ_DOKUMENTACJI := SubStr(NEW_RODZAJ_DOKUMENTACJI, 22) 
			If (NEW_RODZAJ_DOKUMENTACJI = "")
				NEW_RODZAJ_DOKUMENTACJI := " "
		}
		else
			OLD_RODZAJ_DOKUMENTACJI := ""
		
		OLD_NAZWA_TOMU_OBIEKTU := "NEW_NAZWA_TOMU_OBIEKTU"
		FileReadLine, NEW_NAZWA_TOMU_OBIEKTU, %Repozytorium%\%Numer_dok_2%\info.txt, 2
		SPR_NAZWA_TOMU_OBIEKTU := SubStr(NEW_NAZWA_TOMU_OBIEKTU,1,20) 
		SPR2_NAZWA_TOMU_OBIEKTU := "Nazwa Tomu/Obiektu:"A_Tab
		If (SPR_NAZWA_TOMU_OBIEKTU = SPR2_NAZWA_TOMU_OBIEKTU)
		{
			NEW_NAZWA_TOMU_OBIEKTU := SubStr(NEW_NAZWA_TOMU_OBIEKTU, 21) 
			If (NEW_NAZWA_TOMU_OBIEKTU = "")
				NEW_NAZWA_TOMU_OBIEKTU := " "
		}
		else
			OLD_NAZWA_TOMU_OBIEKTU := ""
		
		OLD_NAZWA_PODTOMU_CZESCI := "NEW_NAZWA_PODTOMU_CZESCI"
		FileReadLine, NEW_NAZWA_PODTOMU_CZESCI, %Repozytorium%\%Numer_dok_2%\info.txt, 3
		SPR_NAZWA_PODTOMU_CZESCI := SubStr(NEW_NAZWA_PODTOMU_CZESCI,1,22) 
		SPR2_NAZWA_PODTOMU_CZESCI := "Nazwa Podtomu/części:"A_Tab
		If (SPR_NAZWA_PODTOMU_CZESCI = SPR2_NAZWA_PODTOMU_CZESCI)
		{
			NEW_NAZWA_PODTOMU_CZESCI := SubStr(NEW_NAZWA_PODTOMU_CZESCI, 23) 
			If (NEW_NAZWA_PODTOMU_CZESCI = "")
				NEW_NAZWA_PODTOMU_CZESCI := " "
		}
		else
			OLD_NAZWA_PODTOMU_CZESCI := ""
		
		OLD_PROJEKTANT := "NEW_PROJEKTANT"
		FileReadLine, NEW_PROJEKTANT, %Repozytorium%\%Numer_dok_2%\info.txt, 4
		SPR_PROJEKTANT := SubStr(NEW_PROJEKTANT,1,25) 
		SPR2_PROJEKTANT := "Projektant/Opracowujący:"A_Tab
		If (SPR_PROJEKTANT = SPR2_PROJEKTANT)
		{
			NEW_PROJEKTANT := SubStr(NEW_PROJEKTANT, 26) 
			If (NEW_PROJEKTANT = "")
				NEW_PROJEKTANT := " "
		}
		else
			OLD_PROJEKTANT := ""
		
		OLD_SPRAWDZAJACY := "NEW_SPRAWDZAJACY"
		FileReadLine, NEW_SPRAWDZAJACY, %Repozytorium%\%Numer_dok_2%\info.txt, 5
		SPR_SPRAWDZAJACY := SubStr(NEW_SPRAWDZAJACY,1,15) 
		SPR2_SPRAWDZAJACY := "Sprawdzający:"A_Tab A_tab
		If (SPR_SPRAWDZAJACY = SPR2_SPRAWDZAJACY)
		{
			NEW_SPRAWDZAJACY := SubStr(NEW_SPRAWDZAJACY, 16) 
			If (NEW_SPRAWDZAJACY = "")
				NEW_SPRAWDZAJACY := " "
		}
		else
			OLD_SPRAWDZAJACY := ""
		
; Dane na podstawie Informacji o dokumencie
		
		OLD_BRANZA := "NEW_BRANZA"
		FileReadLine, NEW_BRANZA, %Nowa_dok_info%\info.txt, 9
		
		OLD_NUMER_DOKUMENTU := "NEW_NUMER_DOKUMENTU"
		FileReadLine, NEW_NUMER_DOKUMENTU, %Nowa_dok_info%\info.txt, 2
		
		OLD_WERSJA := "NEW_WERSJA"
		FileReadLine, Rewizja_1, %Nowa_dok_info%\info.txt, 3
		FileReadLine, Rewizja_2, %Nowa_dok_info%\info.txt, 4
		Rewizja_3 := Format("{01:02}", Rewizja_2)	
		NEW_WERSJA =  Wersja: %Rewizja_1%.%Rewizja_3%
		
; Proces podmiany danych
		
		Path = %Nowa_dok%
		oWord := ComObjCreate("Word.Application")
		oWord.Visible := False
		oWord.Documents.Open(Path)
		
		FindAndReplace( oWord, OLD_NUMER_PROJEKTU, NEW_NUMER_PROJEKTU )
		FindAndReplace( oWord, OLD_NAZWA_INWESTYCJI, NEW_NAZWA_INWESTYCJI )
		FindAndReplace( oWord, OLD_ADRES_INWESTYCJI, NEW_ADRES_INWESTYCJI )
		FindAndReplace( oWord, OLD_NAZWA_ZAMAWIAJACEGO, NEW_NAZWA_ZAMAWIAJACEGO )
		FindAndReplace( oWord, OLD_ADRES_ZAMAWIAJACEGO, NEW_ADRES_ZAMAWIAJACEGO )
		FindAndReplace( oWord, OLD_NUMER_UMOWY, NEW_NUMER_UMOWY )
		FindAndReplace( oWord, OLD_KIEROWNIK_PROJEKTU, NEW_KIEROWNIK_PROJEKTU )
		
		FindAndReplace( oWord, OLD_RODZAJ_DOKUMENTACJI, NEW_RODZAJ_DOKUMENTACJI )
		FindAndReplace( oWord, OLD_NAZWA_TOMU_OBIEKTU, NEW_NAZWA_TOMU_OBIEKTU )
		FindAndReplace( oWord, OLD_NAZWA_PODTOMU_CZESCI, NEW_NAZWA_PODTOMU_CZESCI )
		FindAndReplace( oWord, OLD_PROJEKTANT, NEW_PROJEKTANT )
		FindAndReplace( oWord, OLD_SPRAWDZAJACY, NEW_SPRAWDZAJACY )
		
		FindAndReplace( oWord, OLD_BRANZA, NEW_BRANZA )
		FindAndReplace( oWord, OLD_NUMER_DOKUMENTU, NEW_NUMER_DOKUMENTU )
		FindAndReplace( oWord, OLD_WERSJA, NEW_WERSJA )
		oWord.Documents.Save()
		oWord.quit()
	}
	
	SplashTextOff
	MsgBox, 64, Informacja, Utworzono nowy dokument	
}

dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%

Loop, files, %Repozytorium%\%Numer_dok_2%\*%filtr2%*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,dok_2
	FileReadLine, Spr_stat, %A_LoopFileFullPath%\info.txt, 6
	If (filtr3 = "")
		GuiControl, dok: , dok_2, %dok_2%
	If else (Filtr3 = Spr_stat)
		GuiControl, dok: , dok_2, %dok_2%
} 
dok_2 := "Wybierz dokument"

Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 


Status_pro := ""
Status1_pro := ""
Status2_pro := ""
Status3_pro := ""
done = 0
Loop, files, %Repozytorium%\%Numer_dok_2%\*, D
{ 
		FileReadLine, Spr_stat2, %A_LoopFileFullPath%\info.txt, 6
		If (Spr_stat2 = "Wydany")
			Status1_pro := "Wydana"
		If else (Spr_stat2 = "Akceptacja")
			Status2_pro := "Akceptacja"
		If else (Spr_stat2 != "Wydany" and Spr_stat2 != "Akceptacja")
			Status3_pro := "Roboczy"
	} 
	If (Status1_pro != "")
		Status_pro = %Status1_pro%
	If else (Status2_pro != "")
		Status_pro = %Status2_pro%
	If else (Status3_pro != "")
		Status_pro = %Status3_pro%
	
	GuiControl, dok: , Status_pro, %Status_pro%

FileReadLine, Licznik2, %Licznik_dir%\info.txt, 2
Num = %licznik1%%licznik2%
Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
Nowa_dok_info=%Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
Loop,
	If FileExist(Nowa_dok_info)
	{
			Num_count := StrLen(Num)
			Num_count := Format("{01:02}", Num_count)
			Num += 1
			Num_count = {01:%Num_count%}
			Num := Format(Num_count, Num)
			Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
			Nowa_dok_info = %Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
			Nowa_dok = %Roboczy%\%Numer_dok_2%\%Numer_dokumetu%.%dok_import_ext%
}
else
	Break			
GuiControl, stw: , Num, %Num%
GuiControl, imp: , Num, %Num%

return

Stworz_Import:
Gui, imp:Submit, noHide

znaki_spec_spr = %Numer_dokumetu%
Gosub, znaki_spec
If (znak_test = 1)
	return

If (checked_imp = 0)
{
	If (branzadok = "")
	{
		MsgBox, Ostrzeżenie!, Wybierz branżę
		return
	}
	If (rodzajdok = "")
	{
		MsgBox, Ostrzeżenie!, Wybierz rodzaj dokumentu
		return
	}
	Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
	Loop, Read, %Dane%\Temp\info_%A_UserName%.txt
	{
		If (A_Index = 1)
		{
			dok_import2_path = % A_LoopReadLine
			SplashTextOn, 180,40, Proces, Trwa proces...
		}
		If else (A_Index > 1)
		{
			SplitPath, A_LoopReadLine,,,dok_import_ext
			Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
			Nowa_dok_info = %Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
			Nowa_dok = %Roboczy%\%Numer_dok_2%\%Numer_dokumetu%.%dok_import_ext%	
			
			Loop
			{
				If FileExist(Nowa_dok_info)
				{
					Num_count := StrLen(Num)
					Num_count := Format("{01:02}", Num_count)
					Num += 1
					Num_count = {01:%Num_count%}
					Num := Format(Num_count, Num)
					Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
					Nowa_dok_info = %Repozytorium%\%Numer_dok_2%\%Numer_dokumetu%
					Nowa_dok = %Roboczy%\%Numer_dok_2%\%Numer_dokumetu%.%dok_import_ext%
				}
				else
					Break			
			}
			
			FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
			
			FileCopy, %dok_import2_path%\%A_LoopReadLine%, %Nowa_dok%
			
			FileCreateDir, %Nowa_dok_info%
			
			FileDelete, %Nowa_dok_info%\his.txt
			FileAppend, 0.01.001`n%data%`nPlik został utworzony przez: %upr_imie%`n`n, %Nowa_dok_info%\his.txt
			
			FileDelete, %Nowa_dok_info%\info.txt
			FileAppend, %Numer_dokumetu%.%dok_import_ext%`n%Numer_dokumetu%`n0`n01`n001`nRoboczy`n%dok_import_ext%`n%Opis_dok%`n%Lista_branz%`n, %Nowa_dok_info%\info.txt
			
			IF (dok_import_ext = "docx" and checked_imp_kluczowe = 1)
			{
; Dane na podstawie Informacji o projekcie
				
				OLD_NUMER_PROJEKTU := "NEW_NUMER_PROJEKTU"
				FileReadLine, NEW_NUMER_PROJEKTU, %Repozytorium%\info.txt, 1
				SPR_NUMER_PROJEKTU := SubStr(NEW_NUMER_PROJEKTU,1,17) 
				SPR2_NUMER_PROJEKTU := "Numer projektu:"A_Tab A_Tab
				If (SPR_NUMER_PROJEKTU = SPR2_NUMER_PROJEKTU)
				{
					NEW_NUMER_PROJEKTU := SubStr(NEW_NUMER_PROJEKTU, 18) 
					If (NEW_NUMER_PROJEKTU = "")
						NEW_NUMER_PROJEKTU := " "
				}
				else
					OLD_NUMER_PROJEKTU := ""
				
				OLD_NAZWA_INWESTYCJI := "NEW_NAZWA_INWESTYCJI"
				FileReadLine, NEW_NAZWA_INWESTYCJI, %Repozytorium%\info.txt, 2
				SPR_NAZWA_INWESTYCJI := SubStr(NEW_NAZWA_INWESTYCJI,1,19) 
				SPR2_NAZWA_INWESTYCJI := "Nazwa inwestycji:"A_Tab A_Tab
				If (SPR_NAZWA_INWESTYCJI = SPR2_NAZWA_INWESTYCJI)
				{
					NEW_NAZWA_INWESTYCJI := SubStr(NEW_NAZWA_INWESTYCJI, 20) 
					If (NEW_NAZWA_INWESTYCJI = "")
						NEW_NAZWA_INWESTYCJI := " "
				}
				else
					OLD_NAZWA_INWESTYCJI := ""
				
				OLD_ADRES_INWESTYCJI := "NEW_ADRES_INWESTYCJI"
				FileReadLine, NEW_ADRES_INWESTYCJI, %Repozytorium%\info.txt, 3
				SPR_ADRES_INWESTYCJI := SubStr(NEW_ADRES_INWESTYCJI,1, 19) 
				SPR2_ADRES_INWESTYCJI := "Adres inwestycji:"A_Tab A_Tab
				If (SPR_ADRES_INWESTYCJI = SPR2_ADRES_INWESTYCJI)
				{
					NEW_ADRES_INWESTYCJI := SubStr(NEW_ADRES_INWESTYCJI, 20) 
					If (NEW_ADRES_INWESTYCJI = "")
						NEW_ADRES_INWESTYCJI := " "
				}
				else
					OLD_ADRES_INWESTYCJI := ""
				
				OLD_NAZWA_ZAMAWIAJACEGO := "NEW_NAZWA_ZAMAWIAJACEGO"
				FileReadLine, NEW_NAZWA_ZAMAWIAJACEGO, %Repozytorium%\info.txt, 4
				SPR_NAZWA_ZAMAWIAJACEGO := SubStr(NEW_NAZWA_ZAMAWIAJACEGO,1,21) 
				SPR2_NAZWA_ZAMAWIAJACEGO := "Nazwa zamawiającego:"A_Tab
				If (SPR_NAZWA_ZAMAWIAJACEGO = SPR2_NAZWA_ZAMAWIAJACEGO)
				{
					NEW_NAZWA_ZAMAWIAJACEGO := SubStr(NEW_NAZWA_ZAMAWIAJACEGO, 22) 
					If (NEW_NAZWA_ZAMAWIAJACEGO = "")
						NEW_NAZWA_ZAMAWIAJACEGO := " "
				}
				else
					OLD_NAZWA_ZAMAWIAJACEGO := ""
				
				OLD_ADRES_ZAMAWIAJACEGO := "NEW_ADRES_ZAMAWIAJACEGO"
				FileReadLine, NEW_ADRES_ZAMAWIAJACEGO, %Repozytorium%\info.txt, 5
				SPR_ADRES_ZAMAWIAJACEGO := SubStr(NEW_ADRES_ZAMAWIAJACEGO,1,21) 
				SPR2_ADRES_ZAMAWIAJACEGO := "Adres zamawiającego:"A_Tab
				If (SPR_ADRES_ZAMAWIAJACEGO = SPR2_ADRES_ZAMAWIAJACEGO)
				{
					NEW_ADRES_ZAMAWIAJACEGO := SubStr(NEW_ADRES_ZAMAWIAJACEGO, 22) 
					If (NEW_ADRES_ZAMAWIAJACEGO = "")
						NEW_ADRES_ZAMAWIAJACEGO := " "
				}
				else
					OLD_ADRES_ZAMAWIAJACEGO := ""
				
				OLD_NUMER_UMOWY := "NEW_NUMER_UMOWY"
				FileReadLine, NEW_NUMER_UMOWY, %Repozytorium%\info.txt, 6
				SPR_NUMER_UMOWY := SubStr(NEW_NUMER_UMOWY,1,14) 
				SPR2_NUMER_UMOWY := "Numer umowy:"A_Tab A_Tab
				If (SPR_NUMER_UMOWY = SPR2_NUMER_UMOWY)
				{
					NEW_NUMER_UMOWY := SubStr(NEW_NUMER_UMOWY, 15) 
					If (NEW_NUMER_UMOWY = "")
						NEW_NUMER_UMOWY := " "
				}
				else
					OLD_NUMER_UMOWY := ""
				
				OLD_KIEROWNIK_PROJEKTU := "NEW_KIEROWNIK_PROJEKTU"
				FileReadLine, NEW_KIEROWNIK_PROJEKTU, %Repozytorium%\info.txt, 7
				SPR_KIEROWNIK_PROJEKTU := SubStr(NEW_KIEROWNIK_PROJEKTU,1,21) 
				SPR2_KIEROWNIK_PROJEKTU := "Kierownik projektu:"A_Tab A_Tab
				If (SPR_KIEROWNIK_PROJEKTU = SPR2_KIEROWNIK_PROJEKTU)
				{
					NEW_KIEROWNIK_PROJEKTU := SubStr(NEW_KIEROWNIK_PROJEKTU, 22) 
					If (NEW_KIEROWNIK_PROJEKTU = "")
						NEW_KIEROWNIK_PROJEKTU := " "
				}
				else
					OLD_KIEROWNIK_PROJEKTU := ""
				
; Dane na podstawie Informacji o dokumentacji
				
				OLD_RODZAJ_DOKUMENTACJI := "NEW_RODZAJ_DOKUMENTACJI"
				FileReadLine, NEW_RODZAJ_DOKUMENTACJI, %Repozytorium%\%Numer_dok_2%\info.txt, 1
				SPR_RODZAJ_DOKUMENTACJI := SubStr(NEW_RODZAJ_DOKUMENTACJI,1,21) 
				SPR2_RODZAJ_DOKUMENTACJI := "Rodzaj dokumentacji:"A_Tab
				If (SPR_RODZAJ_DOKUMENTACJI = SPR2_RODZAJ_DOKUMENTACJI)
				{
					NEW_RODZAJ_DOKUMENTACJI := SubStr(NEW_RODZAJ_DOKUMENTACJI, 22) 
					If (NEW_RODZAJ_DOKUMENTACJI = "")
						NEW_RODZAJ_DOKUMENTACJI := " "
				}
				else
					OLD_RODZAJ_DOKUMENTACJI := ""
				
				OLD_NAZWA_TOMU_OBIEKTU := "NEW_NAZWA_TOMU_OBIEKTU"
				FileReadLine, NEW_NAZWA_TOMU_OBIEKTU, %Repozytorium%\%Numer_dok_2%\info.txt, 2
				SPR_NAZWA_TOMU_OBIEKTU := SubStr(NEW_NAZWA_TOMU_OBIEKTU,1,20) 
				SPR2_NAZWA_TOMU_OBIEKTU := "Nazwa Tomu/Obiektu:"A_Tab
				If (SPR_NAZWA_TOMU_OBIEKTU = SPR2_NAZWA_TOMU_OBIEKTU)
				{
					NEW_NAZWA_TOMU_OBIEKTU := SubStr(NEW_NAZWA_TOMU_OBIEKTU, 21) 
					If (NEW_NAZWA_TOMU_OBIEKTU = "")
						NEW_NAZWA_TOMU_OBIEKTU := " "
				}
				else
					OLD_NAZWA_TOMU_OBIEKTU := ""
				
				OLD_NAZWA_PODTOMU_CZESCI := "NEW_NAZWA_PODTOMU_CZESCI"
				FileReadLine, NEW_NAZWA_PODTOMU_CZESCI, %Repozytorium%\%Numer_dok_2%\info.txt, 3
				SPR_NAZWA_PODTOMU_CZESCI := SubStr(NEW_NAZWA_PODTOMU_CZESCI,1,22) 
				SPR2_NAZWA_PODTOMU_CZESCI := "Nazwa Podtomu/części:"A_Tab
				If (SPR_NAZWA_PODTOMU_CZESCI = SPR2_NAZWA_PODTOMU_CZESCI)
				{
					NEW_NAZWA_PODTOMU_CZESCI := SubStr(NEW_NAZWA_PODTOMU_CZESCI, 23) 
					If (NEW_NAZWA_PODTOMU_CZESCI = "")
						NEW_NAZWA_PODTOMU_CZESCI := " "
				}
				else
					OLD_NAZWA_PODTOMU_CZESCI := ""
				
				OLD_PROJEKTANT := "NEW_PROJEKTANT"
				FileReadLine, NEW_PROJEKTANT, %Repozytorium%\%Numer_dok_2%\info.txt, 4
				SPR_PROJEKTANT := SubStr(NEW_PROJEKTANT,1,25) 
				SPR2_PROJEKTANT := "Projektant/Opracowujący:"A_Tab
				If (SPR_PROJEKTANT = SPR2_PROJEKTANT)
				{
					NEW_PROJEKTANT := SubStr(NEW_PROJEKTANT, 26) 
					If (NEW_PROJEKTANT = "")
						NEW_PROJEKTANT := " "
				}
				else
					OLD_PROJEKTANT := ""
				
				OLD_SPRAWDZAJACY := "NEW_SPRAWDZAJACY"
				FileReadLine, NEW_SPRAWDZAJACY, %Repozytorium%\%Numer_dok_2%\info.txt, 5
				SPR_SPRAWDZAJACY := SubStr(NEW_SPRAWDZAJACY,1,15) 
				SPR2_SPRAWDZAJACY := "Sprawdzający:"A_Tab A_tab
				If (SPR_SPRAWDZAJACY = SPR2_SPRAWDZAJACY)
				{
					NEW_SPRAWDZAJACY := SubStr(NEW_SPRAWDZAJACY, 16) 
					If (NEW_SPRAWDZAJACY = "")
						NEW_SPRAWDZAJACY := " "
				}
				else
					OLD_SPRAWDZAJACY := ""
				
; Dane na podstawie Informacji o dokumencie
				
				OLD_BRANZA := "NEW_BRANZA"
				FileReadLine, NEW_BRANZA, %Nowa_dok_info%\info.txt, 9
				
				OLD_NUMER_DOKUMENTU := "NEW_NUMER_DOKUMENTU"
				FileReadLine, NEW_NUMER_DOKUMENTU, %Nowa_dok_info%\info.txt, 2
				
				OLD_WERSJA := "NEW_WERSJA"
				FileReadLine, Rewizja_1, %Nowa_dok_info%\info.txt, 3
				FileReadLine, Rewizja_2, %Nowa_dok_info%\info.txt, 4
				Rewizja_3 := Format("{01:02}", Rewizja_2)	
				NEW_WERSJA =  Wersja: %Rewizja_1%.%Rewizja_3%
				
; Proces podmiany danych
				
				Path = %Nowa_dok%
				oWord := ComObjCreate("Word.Application")
				oWord.Visible := False
				oWord.Documents.Open(Path)
				
				FindAndReplace( oWord, OLD_NUMER_PROJEKTU, NEW_NUMER_PROJEKTU )
				FindAndReplace( oWord, OLD_NAZWA_INWESTYCJI, NEW_NAZWA_INWESTYCJI )
				FindAndReplace( oWord, OLD_ADRES_INWESTYCJI, NEW_ADRES_INWESTYCJI )
				FindAndReplace( oWord, OLD_NAZWA_ZAMAWIAJACEGO, NEW_NAZWA_ZAMAWIAJACEGO )
				FindAndReplace( oWord, OLD_ADRES_ZAMAWIAJACEGO, NEW_ADRES_ZAMAWIAJACEGO )
				FindAndReplace( oWord, OLD_NUMER_UMOWY, NEW_NUMER_UMOWY )
				FindAndReplace( oWord, OLD_KIEROWNIK_PROJEKTU, NEW_KIEROWNIK_PROJEKTU )
				
				FindAndReplace( oWord, OLD_RODZAJ_DOKUMENTACJI, NEW_RODZAJ_DOKUMENTACJI )
				FindAndReplace( oWord, OLD_NAZWA_TOMU_OBIEKTU, NEW_NAZWA_TOMU_OBIEKTU )
				FindAndReplace( oWord, OLD_NAZWA_PODTOMU_CZESCI, NEW_NAZWA_PODTOMU_CZESCI )
				FindAndReplace( oWord, OLD_PROJEKTANT, NEW_PROJEKTANT )
				FindAndReplace( oWord, OLD_SPRAWDZAJACY, NEW_SPRAWDZAJACY )
				
				FindAndReplace( oWord, OLD_BRANZA, NEW_BRANZA )
				FindAndReplace( oWord, OLD_NUMER_DOKUMENTU, NEW_NUMER_DOKUMENTU )
				FindAndReplace( oWord, OLD_WERSJA, NEW_WERSJA )
				oWord.Documents.Save()
				oWord.quit()
			}
		}
	}
	SplashTextOff
	Num_count := StrLen(Num)
	Num_count := Format("{01:02}", Num_count)
	Num += 1
	Num_count = {01:%Num_count%}
	Num := Format(Num_count, Num)
	GuiControl, imp: , Num, %Num%	
}
else
{
	Loop, Read, %Dane%\Temp\info_%A_UserName%.txt
	{
		
		If (A_Index = 1)
		{
			dok_import2_path = % A_LoopReadLine
			SplashTextOn, 180,40, Proces, Trwa proces...
		}
		If else (A_Index > 1)
		{
			SplitPath, A_LoopReadLine, ,,dok_import_ext,Numer_imp_dokumetu
			Nowa_dok = %Roboczy%\%Numer_dok_2%\%A_LoopReadLine%	
			Nowa_dok_info = %Repozytorium%\%Numer_dok_2%\%Numer_imp_dokumetu%
			If FileExist(Nowa_dok_info)
			{
				SplashTextOff
				MsgBox,16,Ostrzeżenie!, Plik o numerze %Numer_imp_dokumetu% już istnieje, plik nie zstanie utworzony
				SplashTextOn, 180,40, Proces, Trwa proces...
			}
			FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
			
			FileCopy, %dok_import2_path%\%A_LoopReadLine%, %Nowa_dok%
			
			FileCreateDir, %Nowa_dok_info%
			
			FileDelete, %Nowa_dok_info%\his.txt
			FileAppend, Rewizja 0.01.001`n%data%`nPlik został utworzony przez: %upr_imie%`n`n, %Nowa_dok_info%\his.txt
			
			FileDelete, %Nowa_dok_info%\info.txt
			FileAppend, %A_LoopReadLine%`n%Numer_imp_dokumetu%`n0`n01`n001`nRoboczy`n%dok_import_ext%`n%Lista_branz%`n, %Nowa_dok_info%\info.txt
			
			IF (dok_import_ext = "docx" and checked_imp_kluczowe = 1)
			{
; Dane na podstawie Informacji o projekcie
				
				OLD_NUMER_PROJEKTU := "NEW_NUMER_PROJEKTU"
				FileReadLine, NEW_NUMER_PROJEKTU, %Repozytorium%\info.txt, 1
				SPR_NUMER_PROJEKTU := SubStr(NEW_NUMER_PROJEKTU,1,17) 
				SPR2_NUMER_PROJEKTU := "Numer projektu:"A_Tab A_Tab
				If (SPR_NUMER_PROJEKTU = SPR2_NUMER_PROJEKTU)
				{
					NEW_NUMER_PROJEKTU := SubStr(NEW_NUMER_PROJEKTU, 18) 
					If (NEW_NUMER_PROJEKTU = "")
						NEW_NUMER_PROJEKTU := " "
				}
				else
					OLD_NUMER_PROJEKTU := ""
				
				OLD_NAZWA_INWESTYCJI := "NEW_NAZWA_INWESTYCJI"
				FileReadLine, NEW_NAZWA_INWESTYCJI, %Repozytorium%\info.txt, 2
				SPR_NAZWA_INWESTYCJI := SubStr(NEW_NAZWA_INWESTYCJI,1,19) 
				SPR2_NAZWA_INWESTYCJI := "Nazwa inwestycji:"A_Tab A_Tab
				If (SPR_NAZWA_INWESTYCJI = SPR2_NAZWA_INWESTYCJI)
				{
					NEW_NAZWA_INWESTYCJI := SubStr(NEW_NAZWA_INWESTYCJI, 20) 
					If (NEW_NAZWA_INWESTYCJI = "")
						NEW_NAZWA_INWESTYCJI := " "
				}
				else
					OLD_NAZWA_INWESTYCJI := ""
				
				OLD_ADRES_INWESTYCJI := "NEW_ADRES_INWESTYCJI"
				FileReadLine, NEW_ADRES_INWESTYCJI, %Repozytorium%\info.txt, 3
				SPR_ADRES_INWESTYCJI := SubStr(NEW_ADRES_INWESTYCJI,1, 19) 
				SPR2_ADRES_INWESTYCJI := "Adres inwestycji:"A_Tab A_Tab
				If (SPR_ADRES_INWESTYCJI = SPR2_ADRES_INWESTYCJI)
				{
					NEW_ADRES_INWESTYCJI := SubStr(NEW_ADRES_INWESTYCJI, 20) 
					If (NEW_ADRES_INWESTYCJI = "")
						NEW_ADRES_INWESTYCJI := " "
				}
				else
					OLD_ADRES_INWESTYCJI := ""
				
				OLD_NAZWA_ZAMAWIAJACEGO := "NEW_NAZWA_ZAMAWIAJACEGO"
				FileReadLine, NEW_NAZWA_ZAMAWIAJACEGO, %Repozytorium%\info.txt, 4
				SPR_NAZWA_ZAMAWIAJACEGO := SubStr(NEW_NAZWA_ZAMAWIAJACEGO,1,21) 
				SPR2_NAZWA_ZAMAWIAJACEGO := "Nazwa zamawiającego:"A_Tab
				If (SPR_NAZWA_ZAMAWIAJACEGO = SPR2_NAZWA_ZAMAWIAJACEGO)
				{
					NEW_NAZWA_ZAMAWIAJACEGO := SubStr(NEW_NAZWA_ZAMAWIAJACEGO, 22) 
					If (NEW_NAZWA_ZAMAWIAJACEGO = "")
						NEW_NAZWA_ZAMAWIAJACEGO := " "
				}
				else
					OLD_NAZWA_ZAMAWIAJACEGO := ""
				
				OLD_ADRES_ZAMAWIAJACEGO := "NEW_ADRES_ZAMAWIAJACEGO"
				FileReadLine, NEW_ADRES_ZAMAWIAJACEGO, %Repozytorium%\info.txt, 5
				SPR_ADRES_ZAMAWIAJACEGO := SubStr(NEW_ADRES_ZAMAWIAJACEGO,1,21) 
				SPR2_ADRES_ZAMAWIAJACEGO := "Adres zamawiającego:"A_Tab
				If (SPR_ADRES_ZAMAWIAJACEGO = SPR2_ADRES_ZAMAWIAJACEGO)
				{
					NEW_ADRES_ZAMAWIAJACEGO := SubStr(NEW_ADRES_ZAMAWIAJACEGO, 22) 
					If (NEW_ADRES_ZAMAWIAJACEGO = "")
						NEW_ADRES_ZAMAWIAJACEGO := " "
				}
				else
					OLD_ADRES_ZAMAWIAJACEGO := ""
				
				OLD_NUMER_UMOWY := "NEW_NUMER_UMOWY"
				FileReadLine, NEW_NUMER_UMOWY, %Repozytorium%\info.txt, 6
				SPR_NUMER_UMOWY := SubStr(NEW_NUMER_UMOWY,1,14) 
				SPR2_NUMER_UMOWY := "Numer umowy:"A_Tab A_Tab
				If (SPR_NUMER_UMOWY = SPR2_NUMER_UMOWY)
				{
					NEW_NUMER_UMOWY := SubStr(NEW_NUMER_UMOWY, 15) 
					If (NEW_NUMER_UMOWY = "")
						NEW_NUMER_UMOWY := " "
				}
				else
					OLD_NUMER_UMOWY := ""
				
				OLD_KIEROWNIK_PROJEKTU := "NEW_KIEROWNIK_PROJEKTU"
				FileReadLine, NEW_KIEROWNIK_PROJEKTU, %Repozytorium%\info.txt, 7
				SPR_KIEROWNIK_PROJEKTU := SubStr(NEW_KIEROWNIK_PROJEKTU,1,21) 
				SPR2_KIEROWNIK_PROJEKTU := "Kierownik projektu:"A_Tab A_Tab
				If (SPR_KIEROWNIK_PROJEKTU = SPR2_KIEROWNIK_PROJEKTU)
				{
					NEW_KIEROWNIK_PROJEKTU := SubStr(NEW_KIEROWNIK_PROJEKTU, 22) 
					If (NEW_KIEROWNIK_PROJEKTU = "")
						NEW_KIEROWNIK_PROJEKTU := " "
				}
				else
					OLD_KIEROWNIK_PROJEKTU := ""
				
; Dane na podstawie Informacji o dokumentacji
				
				OLD_RODZAJ_DOKUMENTACJI := "NEW_RODZAJ_DOKUMENTACJI"
				FileReadLine, NEW_RODZAJ_DOKUMENTACJI, %Repozytorium%\%Numer_dok_2%\info.txt, 1
				SPR_RODZAJ_DOKUMENTACJI := SubStr(NEW_RODZAJ_DOKUMENTACJI,1,21) 
				SPR2_RODZAJ_DOKUMENTACJI := "Rodzaj dokumentacji:"A_Tab
				If (SPR_RODZAJ_DOKUMENTACJI = SPR2_RODZAJ_DOKUMENTACJI)
				{
					NEW_RODZAJ_DOKUMENTACJI := SubStr(NEW_RODZAJ_DOKUMENTACJI, 22) 
					If (NEW_RODZAJ_DOKUMENTACJI = "")
						NEW_RODZAJ_DOKUMENTACJI := " "
				}
				else
					OLD_RODZAJ_DOKUMENTACJI := ""
				
				OLD_NAZWA_TOMU_OBIEKTU := "NEW_NAZWA_TOMU_OBIEKTU"
				FileReadLine, NEW_NAZWA_TOMU_OBIEKTU, %Repozytorium%\%Numer_dok_2%\info.txt, 2
				SPR_NAZWA_TOMU_OBIEKTU := SubStr(NEW_NAZWA_TOMU_OBIEKTU,1,20) 
				SPR2_NAZWA_TOMU_OBIEKTU := "Nazwa Tomu/Obiektu:"A_Tab
				If (SPR_NAZWA_TOMU_OBIEKTU = SPR2_NAZWA_TOMU_OBIEKTU)
				{
					NEW_NAZWA_TOMU_OBIEKTU := SubStr(NEW_NAZWA_TOMU_OBIEKTU, 21) 
					If (NEW_NAZWA_TOMU_OBIEKTU = "")
						NEW_NAZWA_TOMU_OBIEKTU := " "
				}
				else
					OLD_NAZWA_TOMU_OBIEKTU := ""
				
				OLD_NAZWA_PODTOMU_CZESCI := "NEW_NAZWA_PODTOMU_CZESCI"
				FileReadLine, NEW_NAZWA_PODTOMU_CZESCI, %Repozytorium%\%Numer_dok_2%\info.txt, 3
				SPR_NAZWA_PODTOMU_CZESCI := SubStr(NEW_NAZWA_PODTOMU_CZESCI,1,22) 
				SPR2_NAZWA_PODTOMU_CZESCI := "Nazwa Podtomu/części:"A_Tab
				If (SPR_NAZWA_PODTOMU_CZESCI = SPR2_NAZWA_PODTOMU_CZESCI)
				{
					NEW_NAZWA_PODTOMU_CZESCI := SubStr(NEW_NAZWA_PODTOMU_CZESCI, 23) 
					If (NEW_NAZWA_PODTOMU_CZESCI = "")
						NEW_NAZWA_PODTOMU_CZESCI := " "
				}
				else
					OLD_NAZWA_PODTOMU_CZESCI := ""
				
				OLD_PROJEKTANT := "NEW_PROJEKTANT"
				FileReadLine, NEW_PROJEKTANT, %Repozytorium%\%Numer_dok_2%\info.txt, 4
				SPR_PROJEKTANT := SubStr(NEW_PROJEKTANT,1,25) 
				SPR2_PROJEKTANT := "Projektant/Opracowujący:"A_Tab
				If (SPR_PROJEKTANT = SPR2_PROJEKTANT)
				{
					NEW_PROJEKTANT := SubStr(NEW_PROJEKTANT, 26) 
					If (NEW_PROJEKTANT = "")
						NEW_PROJEKTANT := " "
				}
				else
					OLD_PROJEKTANT := " "
				
				OLD_SPRAWDZAJACY := "NEW_SPRAWDZAJACY"
				FileReadLine, NEW_SPRAWDZAJACY, %Repozytorium%\%Numer_dok_2%\info.txt, 5
				SPR_SPRAWDZAJACY := SubStr(NEW_SPRAWDZAJACY,1,15) 
				SPR2_SPRAWDZAJACY := "Sprawdzający:"A_Tab A_tab
				If (SPR_SPRAWDZAJACY = SPR2_SPRAWDZAJACY)
				{
					NEW_SPRAWDZAJACY := SubStr(NEW_SPRAWDZAJACY, 16) 
					If (NEW_SPRAWDZAJACY = "")
						NEW_SPRAWDZAJACY := " "
				}
				else
					OLD_SPRAWDZAJACY := ""
				
; Dane na podstawie Informacji o dokumencie
				
				OLD_BRANZA := "NEW_BRANZA"
				FileReadLine, NEW_BRANZA, %Nowa_dok_info%\info.txt, 9
				
				OLD_NUMER_DOKUMENTU := "NEW_NUMER_DOKUMENTU"
				FileReadLine, NEW_NUMER_DOKUMENTU, %Nowa_dok_info%\info.txt, 2
				
				OLD_WERSJA := "NEW_WERSJA"
				FileReadLine, Rewizja_1, %Nowa_dok_info%\info.txt, 3
				FileReadLine, Rewizja_2, %Nowa_dok_info%\info.txt, 4
				Rewizja_3 := Format("{01:02}", Rewizja_2)	
				NEW_WERSJA =  Wersja: %Rewizja_1%.%Rewizja_3%
				
; Proces podmiany danych
				
				Path = %Nowa_dok%
				oWord := ComObjCreate("Word.Application")
				oWord.Visible := false
				oWord.Documents.Open(Path)
				
				FindAndReplace( oWord, OLD_NUMER_PROJEKTU, NEW_NUMER_PROJEKTU )
				FindAndReplace( oWord, OLD_NAZWA_INWESTYCJI, NEW_NAZWA_INWESTYCJI )
				FindAndReplace( oWord, OLD_ADRES_INWESTYCJI, NEW_ADRES_INWESTYCJI )
				FindAndReplace( oWord, OLD_NAZWA_ZAMAWIAJACEGO, NEW_NAZWA_ZAMAWIAJACEGO )
				FindAndReplace( oWord, OLD_ADRES_ZAMAWIAJACEGO, NEW_ADRES_ZAMAWIAJACEGO )
				FindAndReplace( oWord, OLD_NUMER_UMOWY, NEW_NUMER_UMOWY )
				FindAndReplace( oWord, OLD_KIEROWNIK_PROJEKTU, NEW_KIEROWNIK_PROJEKTU )
				
				FindAndReplace( oWord, OLD_RODZAJ_DOKUMENTACJI, NEW_RODZAJ_DOKUMENTACJI )
				FindAndReplace( oWord, OLD_NAZWA_TOMU_OBIEKTU, NEW_NAZWA_TOMU_OBIEKTU )
				FindAndReplace( oWord, OLD_NAZWA_PODTOMU_CZESCI, NEW_NAZWA_PODTOMU_CZESCI )
				FindAndReplace( oWord, OLD_PROJEKTANT, NEW_PROJEKTANT )
				FindAndReplace( oWord, OLD_SPRAWDZAJACY, NEW_SPRAWDZAJACY )
				
				FindAndReplace( oWord, OLD_BRANZA, NEW_BRANZA )
				FindAndReplace( oWord, OLD_NUMER_DOKUMENTU, NEW_NUMER_DOKUMENTU )
				FindAndReplace( oWord, OLD_WERSJA, NEW_WERSJA )
				oWord.Documents.Save()
				oWord.quit()
			}
		}
	}
	SplashTextOff
}
MsgBox, 64, Informacja, Utworzono nowy dokument	

dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%

Loop, files, %Repozytorium%\%Numer_dok_2%\*%filtr2%*, D
{ 
		SplitPath, A_LoopFileFullPath,,,,dok_2
		FileReadLine, Spr_stat, %A_LoopFileFullPath%\info.txt, 6
		If (filtr3 = "")
			GuiControl, dok: , dok_2, %dok_2%
		If else (Filtr3 = Spr_stat)
			GuiControl, dok: , dok_2, %dok_2%
} 
dok_2 := "Wybierz dokument"

Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

Status_pro := ""
Status1_pro := ""
Status2_pro := ""
Status3_pro := ""
done = 0
Loop, files, %Repozytorium%\%Numer_dok_2%\*, D
{ 
		FileReadLine, Spr_stat2, %A_LoopFileFullPath%\info.txt, 6
		If (Spr_stat2 = "Wydany")
			Status1_pro := "Wydana"
		If else (Spr_stat2 = "Akceptacja")
			Status2_pro := "Akceptacja"
		If else (Spr_stat2 != "Wydany" and Spr_stat2 != "Akceptacja")
			Status3_pro := "Roboczy"
	} 
	If (Status1_pro != "")
		Status_pro = %Status1_pro%
	If else (Status2_pro != "")
		Status_pro = %Status2_pro%
	If else (Status3_pro != "")
		Status_pro = %Status3_pro%
	
	GuiControl, dok: , Status_pro, %Status_pro%	

return

Num:
Gui, stw:Submit, noHide
Gui, imp:Submit, noHide

If (checked_imp != 1)
{
		Numer_dokumetu=%Numer_dok_2%%Licznik3%%branzadok%%Licznik3%%rodzajdok%%Licznik3%%Num%
		GuiControl, stw: , Numer_dokumetu, %Numer_dokumetu%
		GuiControl, imp: , Numer_dokumetu, %Numer_dokumetu%
	}
	return
	
	Filtr:
	Gui, Filtr:Destroy
	
	If (Numer_dok_2 = "Wybierz dokumentację")
	{
		MsgBox, 16, Ostrzeżenie, Wybierz dokumentację
		Gui, Filtr:Destroy
		return
	}
	If (filtr = "")	
	{	
		filtr_1 = 1
		filtr_2 = 0
		filtr_3 = 0
		filtr_4 = 0
		filtr_5 = 0
	}
	Gui, dok:Submit, noHide
	
	Gui, Filtr:Font, Bold
	Gui, Filtr:Add, Text, x10 y10 h20 w180 Section, Wybierz status dokumentu:
	Gui, Filtr:Font,
	Gui, Filtr:Add, Radio, xs+10 h20 vfiltr Group checked%filtr_1%, Pomiń status
	Gui, Filtr:Add, Radio, h20 checked%filtr_2%, Roboczy
	Gui, Filtr:Add, Radio, h20 checked%filtr_3%, Sprawdzany
	Gui, Filtr:Add, Radio, h20 checked%filtr_4%, Akceptacja
	Gui, Filtr:Add, Radio, h20 checked%filtr_5%, Odrzucony
	Gui, Filtr:Font, Bold
	Gui, Filtr:Add, Text, y+10 xs h20 w180 Section, Szukaj w nazwie:
	Gui, Filtr:Font,
	Gui, Filtr:Add, Edit, y+5 w180 vfiltr2, %filtr2%
	
	Gui, Filtr:Add, Button, y+10 xs w60 Default gradio_filtr, Ustaw
	Gui, Filtr:Add, Button, x+10 w60 gradio_filtr_clean, Wyczyść
	Gui, Filtr:Show,,Filtr
	
	return
	
	radio_filtr:
	Gui, Filtr:Submit, noHide
	
	If (filtr=1)
	{
		filtr3 := ""
		filtr_1 = 1
		filtr_2 = 0
		filtr_3 = 0
		filtr_4 = 0
		filtr_5 = 0	
	}
	If (filtr=2)
	{
		filtr3 = Roboczy
		filtr_1 = 0
		filtr_2 = 1
		filtr_3 = 0
		filtr_4 = 0
		filtr_5 = 0	
	}
	If (filtr=3)
	{
		filtr3 = Sprawdzany
		filtr_1 = 0
		filtr_2 = 0
		filtr_3 = 1
		filtr_4 = 0
		filtr_5 = 0
	}
	If (filtr=4)
	{
		filtr3 = Akceptacja
		filtr_1 = 0
		filtr_2 = 0
		filtr_3 = 0
		filtr_4 = 1
		filtr_5 = 0
	}
	If (filtr=5)
	{
		filtr3 = Odrzucony
		filtr_1 = 0
		filtr_2 = 0
		filtr_3 = 0
		filtr_4 = 0
		filtr_5 = 1	
	}
	
	If (filtr3 != "")
	{
		If (filtr2 != "")
			filtr_stat = Filtr: %filtr3%/%filtr2%
		if else (filtr2 = "")
			filtr_stat = Filtr: %filtr3%	
	}
	if else (filtr3 = "")
	{
		If (filtr2 != "")
			filtr_stat = Filtr: %filtr2%
		if else (filtr2 = "")
			filtr_stat := ""
	}
	GuiControl, dok: , filtr_stat, %filtr_stat% 	
	
	dok_2 := "Wybierz dokument||"
	GuiControl, dok: , dok_2, |%dok_2% 	
	
	Loop, files, %Repozytorium%\%Numer_dok_2%\*%filtr2%*, D
	{ 
		SplitPath, A_LoopFileFullPath,,,,dok_2
		FileReadLine, Spr_stat, %A_LoopFileFullPath%\info.txt, 6
		If (filtr3 = "")
			GuiControl, dok: , dok_2, %dok_2%
		If else (Filtr3 = Spr_stat)
			GuiControl, dok: , dok_2, %dok_2%
	} 
	dok_2 := "Wybierz dokument"
	
	return
	
	radio_filtr_clean:
	filtr := ""
	filtr2 := ""
	filtr3 := ""
	filtr_stat := ""
	GuiControl, dok: , filtr_stat, %filtr_stat% 	
	dok_2 := "Wybierz dokument||"
	GuiControl, dok: , dok_2, |%dok_2% 
	dok_2 := "Wybierz dokument"
	Loop, files, %Repozytorium%\%Numer_dok_2%\*, D
	{ 
		SplitPath, A_LoopFileFullPath,,,,dok_2
		GuiControl, dok: , dok_2, %dok_2% 	
} 
filtr_1 = 1
filtr_2 = 0
filtr_3 = 0
filtr_4 = 0
filtr_5 = 0	
GuiControl, Filtr:, filtr2, %filtr2%
GuiControl, Filtr:, filtr, %filtr_1%
return
;_______________________________________________________Pliki PDF_____________________________________________

otworz_pdf:
Gui, Dok:Submit, noHide

If (dok_pdf_2 = "Wybierz plik PDF")
{
		MsgBox, 16, Ostrzeżenie, Wybierz plik PDF
		return
	}
	
	Run, %Repozytorium%\%Numer_dok_2%\%dok_2%\%dok_pdf_2%
	return
	
	dodaj_pdf:
	Gui, Dok:Submit, noHide
	
	If (dok_2 = "Wybierz dokument")
	{
		MsgBox, 16, Ostrzeżenie, Wybierz dokument
		return
	}
	
	FileSelectFile, dok_pdf_2,M,,,*.pdf
	Array := StrSplit(dok_pdf_2, "`n")
	
	SplashTextOn, 180,40, Proces, Trwa proces...
	for index, dok_pdf_2 in Array
	{
		if index = 1
			Dir := dok_pdf_2
		else
		{
			file_to_copy = %dir%\%dok_pdf_2%
			FileCopy, %file_to_copy%, %Repozytorium%\%Numer_dok_2%\%dok_2%, 1
		}
	}
	SplashTextOff
	dok_pdf_2 := "Wybierz plik PDF||"
	GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
	
	Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.pdf, R
	{ 
		SplitPath, A_LoopFileFullPath,dok_pdf_2
		GuiControl, dok: , dok_pdf_2, %dok_pdf_2% 		
	}
	return
	
	usun_pdf:
	Gui, Dok:Submit, noHide
	Gui, usun_pdf:Destroy
	
	If (dok_pdf_2 = "Wybierz plik PDF")
	{
		MsgBox, 16, Ostrzeżenie, Wybierz plik PDF
		return
	}
	
	Gui, usun_pdf:Add, Text, x10 y10 h20 w270, Czy na pewno chcesz usunąć zaznaczony plik PDF?
	Gui, usun_pdf:Add, Button, y+10 x70 w60 gUsun_pdf_2, Usuń
	Gui, usun_pdf:Add, Button, x+10 w60 Default gUsun_pdf_3 , Anuluj
	
	Gui, usun_pdf:Show,, Ostrzeżenie!
	return
	
	Usun_pdf_2:
	SplashTextOn, 180,40, Proces, Trwa proces...
	FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\%dok_pdf_2%
	SplashTextOff
	dok_pdf_2 := "Wybierz plik PDF||"
	GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
	
	Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.pdf, R
	{ 
		SplitPath, A_LoopFileFullPath,dok_pdf_2
		GuiControl, dok: , dok_pdf_2, %dok_pdf_2% 		
	}
	Gui, usun_pdf:Destroy
	return
	
	Usun_pdf_3:
	Gui, usun_pdf:Destroy
	return
	
	wymien_pdf:
	Gui, Dok:Submit, noHide
	
	If (dok_pdf_2 = "Wybierz plik PDF")
	{
		MsgBox, 16, Ostrzeżenie, Wybierz plik PDF
		return
	}
	
	FileSelectFile, dok_new_pdf_2,,,,*.pdf
	If !ErrorLevel
	{
		SplashTextOn, 180,40, Proces, Trwa proces...
		FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\%dok_pdf_2%
		FileCopy, %dok_new_pdf_2%, %Repozytorium%\%Numer_dok_2%\%dok_2%\%dok_pdf_2%, 1
		SplashTextOff
	}
	return
	
;_______________________________________________________ETransmit_____________________________________________
	
	zmien_etr:
	Gui, Dok:Submit, noHide
	
	If (dok_2 = "Wybierz dokument")
	{
		MsgBox, 16, Ostrzeżenie, Wybierz dokument
		return
	}
	
	FileSelectFile, etransmit_2,,,,*.rar;*.zip;*.7z
	
	If ErrorLevel
		return
	
	SplashTextOn, 180,40, Proces, Trwa proces...
	Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.rar, R
	{ 
		FileDelete, % A_LoopFileFullPath		
	} 
	Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.zip, R
	{ 
		FileDelete % A_LoopFileFullPath				
	} 
	Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\*.7z, R
	{ 
		FileDelete % A_LoopFileFullPath				
	} 
	
	FileCopy, %etransmit_2%, %Repozytorium%\%Numer_dok_2%\%dok_2%, 1
	SplashTextOff
	FileSetAttrib, +R, %Repozytorium%\%Numer_dok_2%\%dok_2%
	SplitPath, etransmit_2, etransmit_2
	GuiControl, dok: , etransmit_2, %etransmit_2% 
	return
	
;...................................................................funkcje - Informacyjna..............................................................
	
;_______________________________________________________Informacje o projekcie i dokumentacji_____________________________________________
	
	ed_pro:
	Gui, epro:Destroy
	Gui, stw:Submit, noHide
	
	Gui, epro:Add, Edit, y10 x10 w320 h200 vepro, %Info_projekt%
	Gui, epro:Add, Button, y+10 gepro w60 Default, Zmień 
	
	Gui, epro:Show,, Edytuj dane o projekcie
	return
	
	epro:
	Gui, epro:Submit, noHide
	FileDelete, %Repozytorium%\Info.txt
	FileAppend, %epro%, %Repozytorium%\Info.txt
	
	Info_projekt := ""
	GuiControl, dok: , Info_projekt, |%Info_projekt% 
	FileRead, Info_projekt, %Repozytorium%/info.txt
	GuiControl, dok: , Info_projekt, %Info_projekt% 
	
	Gui, epro:Destroy
	return
	
	edytuj_dok:
	Gui, edok:Destroy
	
	If (Numer_dok_2 = "Wybierz dokumentację")
	{
		MsgBox, 16, Ostrzeżenie, Wybierz dokumentację
		Gui, edok:Destroy
		return
	}
	
	Gui, stw:Submit, noHide
	
	If (Numer_dok_2 = "Wybierz dokumentację")
	{
		MsgBox, 16, Ostrzeżenie, Wybierz dokumentację
		Gui, edok:Destroy
		return
}

Gui, edok:Add, Edit, y10 x10 w320 h200 vedok, %Info_dok%
Gui, edok:Add, Button, y+10 gedok w60 Default, Zmień 

Gui, edok:Show,, Edytuj dane o dokumentacji

return

edok:
Gui, edok:Submit, noHide
FileDelete, %Repozytorium%\%Numer_dok_2%\Info.txt
FileAppend, %edok%, %Repozytorium%\%Numer_dok_2%\Info.txt

Info_dok := ""
GuiControl, dok: , Info_dok, |%Info_dok% 
FileRead, Info_dok, %Repozytorium%\%Numer_dok_2%\Info.txt
GuiControl, dok: , Info_dok, %Info_dok% 

Gui, edok:Destroy
return

;_______________________________________________________Informacje o dokumencie_____________________________________________

zmien_dok:
Gui, Stat:Destroy
Gui, dok:Submit, noHide

If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	Gui, Stat:Destroy
	return
}

If (Status_dok = "")
{
	MsgBox 16, Ostrzezenie, Nie wybrano dokumentu
	return
}
If (Status_dok = "Wydany")
{
	Radio_1 = 0
	Radio_2 = 0
	Radio_3 = 0
	Radio_4 = 0	
}
If (Status_dok = "Roboczy")
{
	Radio_1 = 1
	Radio_2 = 0
	Radio_3 = 0
	Radio_4 = 0	
}
If (Status_dok = "Sprawdzany")
{
	Radio_1 = 0
	Radio_2 = 1
	Radio_3 = 0
	Radio_4 = 0	
}
If (Status_dok = "Akceptacja")
{
	Radio_1 = 0
	Radio_2 = 0
	Radio_3 = 1
	Radio_4 = 0	
}
If (Status_dok = "Odrzucony")
{
	Radio_1 = 0
	Radio_2 = 0
	Radio_3 = 0
	Radio_4 = 1	
}

If (upr_uzytkownika = "Pełne" or upr_uzytkownika = "Rozszerzone")
{
	Gui, Stat:Font, Bold
	Gui, Stat:Add, Text, x10 y10 h20 w240 Section, Wybierz aktualny status dokumentu:
	Gui, Stat:Font,
	Gui, Stat:Add, Radio, y+10 xs+10 vRadio_dok Group checked%Radio_1% h20, Roboczy
	Gui, Stat:Add, Radio,h20 checked%Radio_2%, Sprawdzany
	Gui, Stat:Add, Radio,h20 checked%Radio_3%, Akceptacja
	Gui, Stat:Add, Radio,h20 checked%Radio_4% gNotka_stat, Odrzucony
	
	Gui, Stat:Add, Button, y+10 xs w60 Default gchange_stat, Zmień
	Gui, Stat:Show,,Status dokumentu		
}
If else (upr_uzytkownika = "Ograniczone")
{
	Gui, Stat:Font, Bold
	Gui, Stat:Add, Text, x10 y10 h20 w240 Section, Wybierz aktualny status dokumentu:
	Gui, Stat:Font,
	Gui, Stat:Add, Radio, y+10 xs+10 vRadio_dok Group checked%Radio_1% h20, Roboczy
	Gui, Stat:Add, Radio,h20 checked%Radio_2%, Sprawdzany
	Gui, Stat:Add, Radio,h20 checked%Radio_4% gNotka_stat, Odrzucony
	
	Gui, Stat:Add, Button, y+10 xs w60 Default gchange_stat, Zmień
	Gui, Stat:Show,,Status dokumentu	
}

return

Notka_stat:
Gui, stat3:Destroy

not_stat := ""
Gui, stat3:Font, Bold
Gui, stat3:Add, Text, y10 x10 w340 h20, Komentarz do odrzuconego dokumentu:
Gui, stat3:Font,
Gui, stat3:Add, Edit,y+5 w340 h200 vnot_stat, 
Gui, stat3:Add, Button, y+10 w60 gNotka_stat2 Default, Dodaj 

Gui, stat3:Show,, Notka do dokumentu
return

Notka_stat2:
Gui, stat3:Submit,Nohide
Gui, stat3:Destroy
return

stat3GuiClose:
not_stat := ""
Gui, stat3:Destroy
return

zmien2_dok:
Gui, Stat2:Destroy
Gui, dok:Submit, noHide

If (Status_pro = "")
{
	MsgBox 16, Ostrzezenie, Nie wybrano projektu lub w projecie brak dokumentów
	return
}
If (Status_pro = "Wydana")
{
	Radio_5 = 0
	Radio_6 = 0
}
If (Status_pro = "Roboczy")
{
	Radio_5 = 1
	Radio_6 = 0
}
If (Status_pro = "Akceptacja")
{
	Radio_5 = 0
	Radio_6 = 1
}

Gui, Stat2:Font, Bold
Gui, Stat2:Add, Text, x10 y10 h20 w240 Section, Wybierz aktualny status dokumentu:
Gui, Stat2:Font,
Gui, Stat2:Add, Radio, y+10 xs+10 vRadio_pro Group checked%Radio_5% h20, Roboczy
Gui, Stat2:Add, Radio,h20 checked%Radio_6%, Akceptacja

Gui, Stat2:Add, Button, y+10 xs w60 Default gchange2_stat, Zmień
Gui, Stat2:Show,,Status dokumentacji
return

change2_stat:
Gui, Stat2:Submit, 

If (Radio_pro=1)
	Status_pro := "Roboczy"
If (Radio_pro=2)
	Status_pro := "Akceptacja"
GuiControl, dok:, Status_pro, %Status_pro%

Loop, files, %Repozytorium%\%Numer_dok_2%\*, D
{ 
	Org_dok_2 := ""
	Nr_dok_2 := ""
	Rewizja_1 := ""
	Rewizja_2 := ""
	Rewizja_4 := ""
	Status_dok := ""
	dok_2 = %A_LoopFileName%
	FileReadLine, Org_dok_2, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
	FileReadLine, Nr_dok_2, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 2
	FileReadLine, Rewizja_1, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 3
	FileReadLine, Rewizja_2, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 4
	FileReadLine, Rewizja_4, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 5
	FileReadLine, Status_dok, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 6
	
	If (Status_dok = "Wydany")
	{
		If (Status_pro = "Roboczy")
			Rewizja_2++
	}
	Status_dok = %Status_pro%
	
	FileReadLine, dok_ext, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 7
	FileRead, His_dok, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
	
	Rewizja_3 := Format("{01:02}", Rewizja_2)		
	Rewizja_5 := Format("{01:03}", Rewizja_4)
	Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%
	
	FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
	FileAppend, %Rewizja_dok%`n%data%`nZmiana statusu przez %upr_imie% na %Status_dok%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
	FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\Info.txt
	FileAppend, %Org_dok_2%`n%Nr_dok_2%`n%Rewizja_1%`n%Rewizja_3%`n%Rewizja_5%`n%Status_dok%`n%dok_ext%`n, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt
} 

dok_2 := "Wybierz dokument||"
GuiControl, dok: , dok_2, |%dok_2%
Loop, files, %Repozytorium%\%Numer_dok_2%\*%filtr2%*, D
{ 
	SplitPath, A_LoopFileFullPath,,,,dok_2
	FileReadLine, Spr_stat, %A_LoopFileFullPath%\info.txt, 6
	If (filtr3 = "")
		GuiControl, dok: , dok_2, %dok_2%
	If else (Filtr3 = Spr_stat)
		GuiControl, dok: , dok_2, %dok_2%
} 
dok_2 := "Wybierz dokument"

Rewizja_dok := ""
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
Status_dok := ""
GuiControl, dok: , Status_dok, %Status_dok% 	
His_dok := ""
GuiControl, dok: , His_dok, %His_dok% 	
dok_pdf_2 := "Wybierz plik PDF||"
GuiControl, dok: , dok_pdf_2, |%dok_pdf_2% 
etransmit_2 := ""
GuiControl, dok: , etransmit_2, %etransmit_2% 
arch_dok_2 := ""
GuiControl, dok: , arch_dok_2, |%arch_dok_2% 
Opis_dok := ""
GuiControl, dok: , Opis_dok, %Opis_dok% 

Gui, Stat2:Destroy
return

change_stat:
Gui, Stat:Submit,NoHide

FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
Status_sprawdzenie = % Status_dok
If (Radio_dok=1)
{
		Status_dok := "Roboczy"
		not_stat := ""
	}
If (Radio_dok=2)
	{
		Status_dok := "Sprawdzany"
		not_stat := ""
	}
If (Radio_dok=3)
	{
		Status_dok := "Akceptacja"
		not_stat := ""
	}
If (Radio_dok=4)
	{
		Status_dok := "Odrzucony"
		If (not_stat != "")
			not_stat = `nKomentarz do odrzuconego dokumentu:`n%not_stat%
		else
			not_stat := ""
	}

If (Status_sprawdzenie = "Wydany")
{
		If (Status_dok = "Roboczy")
			Rewizja_2++
}

Rewizja_4++
Rewizja_3 := Format("{01:02}", Rewizja_2)		
Rewizja_5 := Format("{01:03}", Rewizja_4)		

Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%

FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileAppend, %Rewizja_dok%`n%data%`nZmiana statusu przez %upr_imie% na %Status_dok%%not_stat%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\Info.txt
FileAppend, %Org_dok_2%`n%Nr_dok_2%`n%Rewizja_1%`n%Rewizja_3%`n%Rewizja_5%`n%Status_dok%`n%dok_ext%`n%Opis_dok%`n, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt

FileReadLine, Org_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
FileReadLine, Rewizja_1 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 3
FileReadLine, Rewizja_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 4
FileReadLine, Rewizja_4 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 5
FileReadLine, Status_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 6
FileReadLine, dok_ext ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 7
FileReadLine, Opis_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 8
FileRead, His_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt

Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%

If (Status_dok = "Akceptacja")
	FileSetAttrib, +R, %Roboczy%\%Numer_dok_2%\%Org_dok_2%
else
	FileSetAttrib, -R, %Roboczy%\%Numer_dok_2%\%Org_dok_2%

GuiControl, dok: , Nr_dok_2, %Nr_dok_2% 
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
GuiControl, dok: , Status_dok, %Status_dok% 	
GuiControl, dok: , His_dok, %His_dok% 	

Status_pro := ""
Status1_pro := ""
Status2_pro := ""
Status3_pro := ""
done = 0
Loop, files, %Repozytorium%\%Numer_dok_2%\*, D
{ 
		FileReadLine, Spr_stat2, %A_LoopFileFullPath%\info.txt, 6
		If (Spr_stat2 = "Wydany")
			Status1_pro := "Wydana"
		If else (Spr_stat2 = "Akceptacja")
			Status2_pro := "Akceptacja"
		If else (Spr_stat2 != "Wydany" and Spr_stat2 != "Akceptacja")
			Status3_pro := "Roboczy"
} 
If (Status1_pro != "")
	Status_pro = %Status1_pro%
If else (Status2_pro != "")
	Status_pro = %Status2_pro%
If else (Status3_pro != "")
	Status_pro = %Status3_pro%

GuiControl, dok: , Status_pro, %Status_pro%	

Gui, Stat:Destroy
return

edytuj_opis_dok:
Gui, Dok:Submit,Nohide
Gui, opis:Destroy

If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	return
}

Gui, opis:Add, Text, y10 x10 Section, Opis dokumentu:
Gui, opis:Add, Edit, y+5 w200 vopis_dok, %Opis_dok%
Gui, opis:Add, Button, y+10 w60 Default gedytuj_opis2_dok, Zmień

Gui, opis:Show, , Opis dokumentu
return

edytuj_opis2_dok:
Gui, opis:Submit,Nohide

FileReadLine, Org_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
FileReadLine, Rewizja_1 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 3
FileReadLine, Rewizja_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 4
FileReadLine, Rewizja_4 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 5
FileReadLine, Status_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 6
FileReadLine, dok_ext ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 7
Rewizja_3 := Format("{01:02}", Rewizja_2)		
Rewizja_5 := Format("{01:03}", Rewizja_4)
Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%

FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\Info.txt
FileAppend, %Org_dok_2%`n%Nr_dok_2%`n%Rewizja_1%`n%Rewizja_3%`n%Rewizja_5%`n%Status_dok%`n%dok_ext%`n%Opis_dok%`n, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt

GuiControl, dok:, Opis_dok, %Opis_dok%

Gui, opis:Destroy
return

zmien_rev:
Gui, Rev:Destroy
Gui, Dok:Submit, noHide

If (dok_2 = "Wybierz dokument")
{
		MsgBox, 16, Ostrzeżenie, Wybierz dokument
		return
}

FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss

old_Rewizja_1 = %Rewizja_1%
old_Rewizja_2 = %Rewizja_2%
old_Rewizja_3 := Format("{01:02}", Old_Rewizja_2)
OLD_WERSJA = Wersja: %old_Rewizja_1%.%old_Rewizja_2%

Gui, Rev:Add, Text, y10 x10 Section w150, Wprowadź rewizję główną:
Gui, Rev:Add, Button, x+10 Center grev_1_plus w20, +
Gui, Rev:Add, Edit,x+5 w60 vgl_rev, %Rewizja_1%
Gui, Rev:Add, Button, x+5 Center w20 grev_1_minus, -
Gui, Rev:Add, Text,y+10 w150 xs, Wprowadź rewizję pomocniczą:
Gui, Rev:Add, Button, x+10 Center w20 grev_2_plus, +
Gui, Rev:Add, Edit,x+5 w60 vpo_rev, %Rewizja_2%
Gui, Rev:Add, Button, x+5 Center w20 grev_2_minus, -
Gui, Rev:Add, Checkbox, y+10 xs w200 vchecked_rev_klucze, Zmień wewnątrz dokumentu „docx”
Gui, Rev:Add, Button, y+10 xs w60 Default gchange_rev, Zmień

Gui, Rev:Show,,Rewizja dokumentu
return

rev_1_plus:
Gui, Rev:Submit, noHide
gl_rev += 1
po_rev = 0
po_rev := Format("{01:02}", po_rev)
GuiControl, Rev:, po_rev, %po_rev%
GuiControl, Rev:, gl_rev, %gl_rev%
return

rev_1_minus:
Gui, Rev:Submit, noHide
If (gl_rev > "0")
{
	gl_rev -= 1
	po_rev = 0
	po_rev := Format("{01:02}", po_rev)
	GuiControl, Rev:, po_rev, %po_rev%
	GuiControl, Rev:, gl_rev, %gl_rev%	
}
return

rev_2_plus:
Gui, Rev:Submit, noHide
po_rev += 1
po_rev := Format("{01:02}", po_rev)
GuiControl, Rev:, po_rev, %po_rev%
return

rev_2_minus:
Gui, Rev:Submit, noHide
If (po_rev > "00")
{
	po_rev -= 1
	po_rev := Format("{01:02}", po_rev)
	GuiControl, Rev:, po_rev, %po_rev%
	GuiControl, Rev:, po_rev, %po_rev%	
}
return

change_rev:
Gui, Rev:Submit

Rewizja_1=%gl_rev%
Rewizja_2=%po_rev%
Rewizja_3 := Format("{01:02}", Rewizja_2)	
Rewizja_4++
Rewizja_5 := Format("{01:03}", Rewizja_4)

Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%

FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileAppend, %Rewizja_dok%`n%data%`nZmiana rewizji przez %upr_imie% na %Status_dok%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\Info.txt
FileAppend, %Org_dok_2%`n%Nr_dok_2%`n%Rewizja_1%`n%Rewizja_3%`n%Rewizja_5%`n%Status_dok%`n%dok_ext%`n, %Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt

FileReadLine, Org_dok_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
FileReadLine, Rewizja_1 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 3
FileReadLine, Rewizja_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 4
FileReadLine, Rewizja_4 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 5
FileReadLine, Status_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 6
FileReadLine, dok_ext ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 7
FileRead, His_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt

Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%

GuiControl, dok: , Nr_dok_2, %Nr_dok_2% 
GuiControl, dok: , Rewizja_dok, %Rewizja_dok% 	
GuiControl, dok: , Status_dok, %Status_dok% 	
GuiControl, dok: , His_dok, %His_dok%

If (dok_ext = "docx" and checked_rev_klucze = 1)
{
	SplashTextOn, 180,40, Proces, Trwa proces...
	NEW_WERSJA = Wersja: %Rewizja_1%.%Rewizja_3%
	Path = %Roboczy%\%Numer_dok_2%\%Org_dok_2%
	oWord := ComObjCreate("Word.Application")
	oWord.Visible := false
	oWord.Documents.Open(Path)
	FindAndReplace( oWord, OLD_WERSJA, NEW_WERSJA )
	oWord.Documents.Save()
	oWord.quit()	
	SplashTextOff
}

Gui, Rev:Destroy
return

Notka_dok:
Gui, enot:Destroy
Gui, dok:Submit, noHide

If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	Gui, enot:Destroy
	return
}

Gui, enot:Add, Edit,y15 w340 h200 venot, 
Gui, enot:Add, Button, y+10 genot w60 Default, Dodaj 

Gui, enot:Show,, Notka do dokumentu

return

enot:
Gui, enot:Submit, noHide
FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileAppend, %Rewizja_1%_%Rewizja_3%`n%data%`nDodano notke sporządzoną przez %upr_imie%`n%enot%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt

Info_dok := ""
GuiControl, dok: , His_dok, |%His_dok% 
FileRead, His_dok, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
GuiControl, dok: , His_dok, %His_dok% 

Gui, enot:Destroy
return

zalacznik:
Gui, dok:Submit, noHide
Gui, zal:Destroy
If (dok_2 = "Wybierz dokument")
{
	MsgBox, 16, Ostrzeżenie, Wybierz dokument
	Gui, zal:Destroy
	return
}

his_dr = %Repozytorium%\%Numer_dok_2%\%dok_2%\his
If !FileExist(his_dr)
	FileCreateDir, %his_dr%

Gui, zal:Font, Bold
Gui, zal:Add, Text, x10 y10, Pliki będące załącznikiem do wpisu w historii dokumentu: 
Gui, zal:Font, 
Gui, zal:Add, ListBox, y+10 w340 h200 vzal_note_2, %zal_note_2%
Gui, zal:Add, Button, y+10 w60 Default gOtworz_zal, Otwórz
Gui, zal:Add, Button, x+10 w60 gDodaj_zal, Dodaj

zal_note_2 := ""
GuiControl, zal: , zal_note_2, |%zal_note_2%
Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\his\*, R
{ 
	SplitPath, A_LoopFileFullPath,zal_note_2
	GuiControl, zal: , zal_note_2, %zal_note_2% 		
} 

Gui, zal:Show,, Załączone notki
return

Otworz_zal:
Gui, zal:Submit, noHide
If (zal_note_2 != "")
	Run, %Repozytorium%\%Numer_dok_2%\%dok_2%\his\%zal_note_2%
else
	MsgBox, 16, Ostrzeżenie, Wybierz załącznik
return

Dodaj_zal:
Gui, zal:Submit, noHide
FileSelectFile, zal_note
SplitPath, zal_note,zal_note_2
zal_note_3 = %Repozytorium%\%Numer_dok_2%\%dok_2%\his\%zal_note_2%
FormatTime, data_2, A_Now, yyyy_MM_dd_HHmmss
SplashTextOn, 180,40, Proces, Trwa proces...
FileCopy, %zal_note%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his\%data_2%_%zal_note_2%, 1
SplashTextOff
FileSetAttrib, +R, %Repozytorium%\%Numer_dok_2%\%dok_2%\his\%data_2%_%zal_note_2%

FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileAppend, %Rewizja_1%.%Rewizja_3%.%Rewizja_5%`n%data%`nDodano załącznik %data_2%_%zal_note_2% do notki przez %upr_imie%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileRead, His_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
GuiControl, dok: , His_dok, %His_dok% 	

zal_note_2 := ""
GuiControl, zal: , zal_note_2, |%zal_note_2%
Loop, files, %Repozytorium%\%Numer_dok_2%\%dok_2%\his\*, R
{ 
	SplitPath, A_LoopFileFullPath,zal_note_2
	GuiControl, zal: , zal_note_2, %zal_note_2%
}		
return

Przywroc_arch:
Gui, dok:Submit, noHide

If (his_dok = "")
{
	MsgBox, 16, Ostrzeżenie, Wybierz archiwum
	return
}

arch_dok_3 = % arch_dok_2
Gosub, Archiwizuj
SplashTextOn, 180,40, Proces, Trwa proces...

Rewizja_1 := SubStr(arch_dok_3,20,1) 
Rewizja_2 := SubStr(arch_dok_3,22,2) 
Rewizja_3 := Format("{01:02}", Rewizja_2)
Rewizja_dok =  %Rewizja_1%.%Rewizja_3%.%Rewizja_5%

FileCopy, %Repozytorium%\%Numer_dok_2%\%dok_2%\%arch_dok_3%\%dok_org_2%, %Roboczy%\%Numer_dok_2%\%dok_org_2%, 1
SplashTextOff
FileSetAttrib, -R, %Roboczy%\%Numer_dok_2%\%dok_org_2%
FormatTime, data, A_Now, yyyy_MM_dd_HH:mm:ss
FileDelete, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileAppend, %Rewizja_1%.%Rewizja_3%.%Rewizja_5%`n%data%`nPrzywrócono wersję archiwalną przez %upr_imie%`n`n%his_dok%, %Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
FileRead, His_dok ,%Repozytorium%\%Numer_dok_2%\%dok_2%\his.txt
GuiControl, dok: , His_dok, %His_dok% 	
GuiControl, dok: , Rewizja_dok, %Rewizja_dok%

MsgBox, 64, Informacja, Plik został przywrócony z archiwum
return

Otworz_arch:
Gui, dok:Submit, noHide

If (his_dok = "")
{
	MsgBox, 16, Ostrzeżenie, Wybierz archiwum
	return
}

FileReadLine, dok_org_2 ,%Repozytorium%\%Numer_dok_2%\%dok_2%\info.txt, 1
Run, %Repozytorium%\%Numer_dok_2%\%dok_2%\%arch_dok_2%\%dok_org_2%
return

;________________________________________________________________________Zamykanie okien______________________________

dokGuiClose:
Gui, Zamknij:Destroy

Gui, Zamknij:Add, Text, x10 y10 w270, Czy na pewno chcesz zamknąć program?
Gui, Zamknij:Add, Button, y+10 xp+70 w60 gzamknij1, Tak
Gui, Zamknij:Add, Button, x+10 w60 Default gzamknij2 , Anuluj

Gui, Zamknij:Show,, Ostrzeżenie!
return

zamknij1:
ExitApp

zamknij2:
Gui, Zamknij:Destroy
return

nproGuiClose:
Gui, R_gen:Destroy
Gui, Gen:Destroy
Gui, npro:Destroy
return

GenGuiClose:
Gui, R_gen:Destroy
Gui, Gen:Destroy
return

przelaczGuiClose:
Gui, dodaj2:Destroy
Gui, dodaj:Destroy
Gui, przelacz:Destroy
return

stwGuiClose:
Gui, S1:Destroy
Gui, R1_rodz:Destroy
Gui, R1_br:Destroy
Gui, dod_branza:Destroy
Gui, dod_2_branza:Destroy
Gui, Rodz2_dok:Destroy
Gui, Rodz_dok:Destroy
Gui, Num1:Destroy
Gui, Num2:Destroy
Gui, Num3:Destroy
Gui, Num4:Destroy
Gui, Num5:Destroy
Gui, stw:Destroy
return

impGuiClose:
Gui, S1:Destroy
Gui, R1_rodz:Destroy
Gui, R1_br:Destroy
Gui, dod_branza:Destroy
Gui, dod_2_branza:Destroy
Gui, Rodz2_dok:Destroy
Gui, Rodz_dok:Destroy
Gui, Num1:Destroy
Gui, Num2:Destroy
Gui, Num3:Destroy
Gui, Num4:Destroy
Gui, Num5:Destroy
Gui, imp:Destroy
return

Num2GuiClose:
Gui, Num3:Destroy
Gui, Num4:Destroy
Gui, Num5:Destroy
Gui, Num2:Destroy
return

dodajGuiClose:
If !FileExist(info_DMS)
	ExitApp
Gui, dodaj:Destroy
return

uprawnieniaGuiClose:
Gui, uprawnienia5:Destroy
Gui, uprawnienia4:Destroy
Gui, uprawnienia2:Destroy
Gui, uprawnienia:Destroy
return

uprawnienia2GuiClose:
Gui, uprawnienia5:Destroy
Gui, uprawnienia4:Destroy
Gui, uprawnienia2:Destroy
return
;________________________________________________________________________Instrukcja______________________________

labelToolTips:

MouseGetPos,,, TheWin , TheControl

IfWinActive, %Tytul_projektu%
{
	
	If (TheControl="Button1")
		ToolTip, Załóż nową dokumentację
	else If (TheControl="Button2")
		ToolTip, Edytuj numer zaznaczonej dokumentacji
	else If (TheControl="Button3")
		ToolTip, Usuń zaznaczoną dokumentację
	else If (TheControl="Button4")
		ToolTip, Wydaj zaznaczoną dokumentację do folderu Dokumentacja
	else If (TheControl="Button5")
		ToolTip, Utwórz roaport dokumentów dla zaznaczonej dokumentacji
	else If (TheControl="Button6")
		ToolTip, Odśwież wykaz dostępnych dokumentacji
	else If (TheControl="Button7")
		ToolTip, Przełącz/edytuj/dodaj projekt
	else If (TheControl="Button8")
		ToolTip, Ustal uprawnienia dla projektu
	else If (TheControl="Button9")
		ToolTip, Informacje o programie
	else If (TheControl="Button10")
		ToolTip, Otwórz zaznaczony dokument
	else If (TheControl="Button11")
		ToolTip, Otwórz w trybie do odczytu zaznaczony dokument
	else If (TheControl="Button12")
		ToolTip, Utwórz nowy dokument na podstawie szablonu
	else If (TheControl="Button13")
		ToolTip, Utwórz nowy dokument na podstawie importowanego pliku
	else If (TheControl="Button14")
		ToolTip, Edytuj nazwę dokumentu
	else If (TheControl="Button15")
		ToolTip, Wymień zaznaczony dokument
	else If (TheControl="Button16")
		ToolTip, Usuń zaznaczony dokument
	else If (TheControl="Button17")
		ToolTip, Utwórz archiwum dla dokumentu
	else If (TheControl="Button18")
		ToolTip, Filtrój dokumenty po statusie i nazwie
	else If (TheControl="Button19")
		ToolTip, Edytuj informacje na temat projektu
	else If (TheControl="Button20")
		ToolTip, Edytuj informacje na temat dokumentacji
	else If (TheControl="Button22")
		ToolTip, Zmień status dla całej dokumentacji
	else If (TheControl="Button24")
		ToolTip, Otwórz zaznaczony plik PDF
	else If (TheControl="Button25")
		ToolTip, Dodaj plik PDF do zaznaczonego dokumentu
	else If (TheControl="Button26")
		ToolTip, Usuń zaznaczony plik PDF
	else If (TheControl="Button27")
		ToolTip, Wymień zaznaczony plik PDF
	else If (TheControl="Button28")
		ToolTip, Zmień\Dodaj plik ETransmit dla zaznaczonego dokumentu (rar/zip/7z)
	else If (TheControl="Button29")
		ToolTip, Otwórz zaznaczone archiwum
	else If (TheControl="Button30")
		ToolTip, Przwróc zanzaczone archiwum jako dokument
	else If (TheControl="Button31")
		ToolTip, Zmień opis zaznaczonego dokumentu
	else If (TheControl="Button32")
		ToolTip, Zmień ręcznie rewizję zaznaczonego dokumentu
	else If (TheControl="Button33")
		ToolTip, Zmień satus zaznaczonego dokumentu
	else If (TheControl="Button34")
		ToolTip, Dodaj notatkę do historii dokumentu
	else If (TheControl="Button35")
		ToolTip, Załącz plik z notatką
	else If (TheControl="ListBox1")
		ToolTip, Dostępne dokumentacje w projekcie
	else If (TheControl="ListBox2")
		ToolTip, Dostępne dokumenty w ramach wybranej dokumentacji
	else If (TheControl="ListBox3")
		ToolTip, Dostępne pliki PDF dołączone do wybranego dokumentu
	else If (TheControl="ListBox4")
		ToolTip, Utworzone archiwa dla wybranego dokumentu
	else 
		ToolTip,
}
else IfWinActive, Tworzenie nowej dokumentacji
{
	If (TheControl="Button1")
		ToolTip, Generuj nazwę dokumentacji ze szablonu
	else If (TheControl="Button2")
		ToolTip, Generuj opis dokumentacji na podstawie słów kluczy
	else If (TheControl="Button3")
		ToolTip, Dodaj nową dokumentację
	else
		ToolTip,
}
else IfWinActive, Edycja nazwy
{
	If (TheControl="Button1")
		ToolTip, Zmień nazwę dokumentacji
	else
		ToolTip,
}
else IfWinActive, Raport
{
	If (TheControl="Button1")
		ToolTip, Exportuj raport do pliku tekstowego
	else If (TheControl="Button2")
		ToolTip, Zamknij okno raportu
	else
		ToolTip,
}
else IfWinActive, Przełącz projekt
{
	If (TheControl="Button1")
		ToolTip, Zmień folder w którym znajdują się projekty
	else If (TheControl="Button2")
		ToolTip, Przełącz na wybrany projekt
	else If (TheControl="Button3")
		ToolTip, Zmień nazwę wybranego projektu
	else If (TheControl="Button4")
		ToolTip, Dodaj nowy projekt
	else If (TheControl="Button5")
		ToolTip, Zamknij okno przełączania projektów
	else
		ToolTip,
}
else IfWinActive, Uprawnienia
{
	If (TheControl="Button1")
		ToolTip, Uprawnienia do pełnej funkcjonalności programu
	else If (TheControl="Button2")
		ToolTip, Uprawnienia ograniczone w sferze tworzenia i edycji projektu
	else If (TheControl="Button3")
		ToolTip, Uprawnienia ograniczone do tworzenia i edycji dokumentów
	else If (TheControl="Button5")
		ToolTip, Zmień/dodaj/usuń uprawnienia dla grup uprawnień
	else
		ToolTip,
}
else IfWinActive, Informacje o programie
{
	If (TheControl="Button1")
		ToolTip, Prześlij uwagi i propozycje za pośrednictwem programem Outlook
	else
		ToolTip,
}
else IfWinActive, Stwórz nowy dokument
{
	If (TheControl="Button1")
		ToolTip, Dodaj nową branżę
	else If (TheControl="Button2")
		ToolTip, Edytuj daną branżę
	else If (TheControl="Button3")
		ToolTip, Usuń daną branżę
	else If (TheControl="Button4")
		ToolTip, Dodaj rodzaj dokumentu
	else If (TheControl="Button5")
		ToolTip, Edytuj dany rodzaj dokumentu
	else If (TheControl="Button6")
		ToolTip, Usuń dany rodzaj dokumentu
	else If (TheControl="Button7")
		ToolTip, Dodaj szablon do bazy
	else If (TheControl="Button8")
		ToolTip, Wymień zaznaczony szablon
	else If (TheControl="Button9")
		ToolTip, Usuń zaznaczony szablon
	else If (TheControl="Button10")
		ToolTip, Ustaw specjalizacje dokumentu
	else If (TheControl="Button11")
		ToolTip, Edytuj sposób numeracji dokumentu
	else If (TheControl="Button12")
		ToolTip, Stwórz nowy dokument
	else If (TheControl="Button13")
		ToolTip, Zamknij okno tworzenia nowych dokumentów
	else
		ToolTip,
}
else IfWinActive, Zaimportuj nowy dokument
{
	If (TheControl="Button1")
		ToolTip, Dodaj nową branżę
	else If (TheControl="Button2")
		ToolTip, Edytuj daną branżę
	else If (TheControl="Button3")
		ToolTip, Usuń daną branżę
	else If (TheControl="Button4")
		ToolTip, Dodaj rodzaj dokumentu
	else If (TheControl="Button5")
		ToolTip, Edytuj dany rodzaj dokumentu
	else If (TheControl="Button6")
		ToolTip, Usuń dany rodzaj dokumentu
	else If (TheControl="Button7")
		ToolTip, Zmień pliki do importu
	else If (TheControl="Button8")
		ToolTip, Ustaw specjalizacje dokumentu
	else If (TheControl="Button9")
		ToolTip, Edytuj sposób numeracji dokumentu
	else If (TheControl="Button11")
		ToolTip, Stwórz nowy dokument z zaimportowanego pliku
	else If (TheControl="Button12")
		ToolTip, Zamknij okno tworzenia nowych dokumentów
	else
		ToolTip,
}
else IfWinActive, Edytuj dokument
{
	If (TheControl="Button1")
		ToolTip, Edytuj nazwę dokumentu
	else
		ToolTip,
}
else IfWinActive, Filtr
{
	If (TheControl="Button6")
		ToolTip, Ustaw filtr
	else If (TheControl="Button7")
		ToolTip, Wyczyść zadany filtr
	else
		ToolTip,
}
else IfWinActive, Edytuj dane o projekcie
{
	If (TheControl="Button1")
		ToolTip, Zmień dane o projekcie
	else
		ToolTip,
}
else IfWinActive, Edytuj dane o dokumentacji
{
	If (TheControl="Button1")
		ToolTip, Zmień dane o dokumentacji
	else
		ToolTip,
}
else IfWinActive, Status dokumentacji
{
	If (TheControl="Button1")
		ToolTip, Dokumentacja w trakcie opracowywania
	else If (TheControl="Button2")
		ToolTip, Dokumentacja zaakceptowany – gotowy do wydania
	else If (TheControl="Button3")
		ToolTip, Zmień status dokumentacji wraz z dokumentami przypisanymi
	else
		ToolTip,
}
else IfWinActive, Opis dokumentu
{
	If (TheControl="Button1")
		ToolTip, Zmień opis zaznaczonego dokumentu
	else
		ToolTip,
}
else IfWinActive, Rewizja dokumentu
{
	If (TheControl="Button5")
		ToolTip, Zmień rewizję zaznaczonego dokumentu
	else
		ToolTip,
}
else IfWinActive, Status dokumentu
{
	If (TheControl="Button1")
		ToolTip, Dokument w trakcie opracowywania
	else If (TheControl="Button2")
		ToolTip, Dokument w trakcie sprawdzania
	else If (TheControl="Button3")
		ToolTip, Dokument zaakceptowany – gotowy do wydania
	else If (TheControl="Button4")
		ToolTip, Dokument odrzucony po sprawdzeniu
	else If (TheControl="Button5")
		ToolTip, Zmień status zaznaczonego dokumentu
	else
		ToolTip,
}
else IfWinActive, Notka do dokumentu
{
	If (TheControl="Button1")
		ToolTip, Dodaj notkę do historii dokumentu
	else
		ToolTip,
}
else IfWinActive, Załączone notki
{
	If (TheControl="Button1")
		ToolTip, Otwórz plik z notką
	else If (TheControl="Button2")
		ToolTip, Dodaj plik z notką
	else
		ToolTip,
}
else IfWinActive, Generuj nazwę dokumentacji
{
	If (TheControl="Button1")
		ToolTip, Baza rodzajów dokumentacji
	else If (TheControl="Button2")
		ToolTip, Generuj nazwę dokumentacji
	else
		ToolTip,
}
else IfWinActive, Rodzaj dokumentacji
{
	If (TheControl="Button1")
		ToolTip, Dodaj rodzaj dokumentacji
	else If (TheControl="Button2")
		ToolTip, Edytuj wybrany rodzaj dokumentacji
	else If (TheControl="Button3")
		ToolTip, Usuń wybrany rodzaj dokumentacji
	else If (TheControl="Button4")
		ToolTip, Wstaw rodzaj dokumentacji
	else
		ToolTip,
}
else IfWinActive, Dodaj rodzaj dokumentacji
{
	If (TheControl="Button1")
		ToolTip, Dodaj rodzaj dokumentacji
	else
		ToolTip,
}
else IfWinActive, Edytuj rodzaj dokumentacji
{
	If (TheControl="Button1")
		ToolTip, Edytuj` rodzaj dokumentacji
	else
		ToolTip,
}
else IfWinActive, Edytuj projekt
{
	If (TheControl="Button2")
		ToolTip, Zmień dane dla zaznaczonego projektu
	else
		ToolTip,
}
else IfWinActive, Tworzenie nowego projektu
{
	If (TheControl="Button2")
		ToolTip, Generuj opis projektu na podstawie słów kluczy
	else If (TheControl="Button3")
		ToolTip, Utwórz nowy projekt
	else
		ToolTip, 
}
else IfWinActive, Dodaj branżę
{
	If (TheControl="Button1")
		ToolTip, Dodaj nową branżę
	else
		ToolTip,
}
else IfWinActive, Edytuj branżę
{
	If (TheControl="Button1")
		ToolTip, Edytuj wybraną branżę
	else
		ToolTip,
}
else IfWinActive, Dodaj rodzaj dokumentu
{
	If (TheControl="Button1")
		ToolTip, Dodaj nowy rodzaj dokumentu
	else
		ToolTip,
}
else IfWinActive, Edytuj rodzaj dokumentu
{
	If (TheControl="Button1")
		ToolTip, Edytuj rodzaj dokumentu 
	else
		ToolTip,
}
else IfWinActive, Specjalizacja dokumentu
{
	If (TheControl="Button1")
		ToolTip, Ustaw specjalizację dokumentu
	else If (TheControl="Button2")
		ToolTip, Dodaj specjalizację dokumentu
	else If (TheControl="Button3")
		ToolTip, Edytuj wybraną specjalizację dokumentu
	else If (TheControl="Button4")
		ToolTip, Usuń wybraną specjalizację dokumentu
	else
		ToolTip,
}
else IfWinActive, Dodaj specjalizację
{
	If (TheControl="Button1")
		ToolTip, Dodaj specjalizację dokumentu
	else
		ToolTip,
}
else IfWinActive, Edytuj specjalizację
{
	If (TheControl="Button1")
		ToolTip, Zmień specjalizację dokumentu
	else
		ToolTip,
}
else IfWinActive, Sposób numeracji dokumentu
{
	If (TheControl="Button1")
		ToolTip, Zmień sposób numeracji dokumentu
	else
		ToolTip,
}
else IfWinActive, Generuj klucze dokumentacji
{
	If (TheControl="Button1")
		ToolTip, Generuj opis
	else
		ToolTip,
}
else IfWinActive, Generuj klucze projektu
{
	If (TheControl="Button1")
		ToolTip, Generuj opis
	else
		ToolTip,
}
else
	ToolTip,
return

;_________________________________________________________Starter pack___________________________________________
starter_dane:
FileAppend,AKP,%Dane2%\Branza\AKPiA.txt
FileAppend,ARC,%Dane2%\Branza\Architektura.txt
FileAppend,BIM,%Dane2%\Branza\Model złożeniowy.txt
FileAppend,DPP,%Dane2%\Branza\Planowanie przestrzenne.txt
FileAppend,DRG,%Dane2%\Branza\Drogowa.txt
FileAppend,EKN,%Dane2%\Branza\Ekonomiczna.txt
FileAppend,ELK,%Dane2%\Branza\Elektryczna.txt
FileAppend,GED,%Dane2%\Branza\Geodezja.txt
FileAppend,GEL,%Dane2%\Branza\Geologia.txt
FileAppend,GET,%Dane2%\Branza\Geotechnika.txt
FileAppend,GIS,%Dane2%\Branza\GIS.txt
FileAppend,HYD,%Dane2%\Branza\Hydraulika.txt
FileAppend,INS,%Dane2%\Branza\Instalacyjna.txt
FileAppend,KAT,%Dane2%\Branza\Ochrona Katodowa.txt
FileAppend,KON,%Dane2%\Branza\Konstrukcja.txt
FileAppend,KST,%Dane2%\Branza\Kosztorysowa.txt
FileAppend,LIN,%Dane2%\Branza\Liniowa.txt
FileAppend,MEL,%Dane2%\Branza\Melioracja.txt
FileAppend,OWY,%Dane2%\Branza\Obliczenia wytrzymałościowe.txt
FileAppend,POZ,%Dane2%\Branza\Ppoż.txt
FileAppend,PRC,%Dane2%\Branza\Procesowa.txt
FileAppend,SIK,%Dane2%\Branza\Studia i Koncepcje.txt
FileAppend,SRD,%Dane2%\Branza\Środowisko.txt
FileAppend,TEC,%Dane2%\Branza\Technologia.txt
FileAppend,TEL,%Dane2%\Branza\Telekomunikacja.txt
FileAppend,WEG,%Dane2%\Branza\Wypisy z ewidencji gruntów.txt
FileAppend,ZAR,%Dane2%\Branza\Zabytki archeologiczne.txt
FileAppend,ZPR,%Dane2%\Branza\Zarządzanie projektem.txt
FileAppend,ZZZ,%Dane2%\Branza\Wielobranżowa.txt

FileAppend,AS0`nPodstawowe,%Dane2%\Rodzaje_dok\Aspekty środowiskowe.txt
FileAppend,BOZ`nPodstawowe,%Dane2%\Rodzaje_dok\Plan bezpieczeństwa i ochrony zdrowia.txt
FileAppend,DBP`nPodstawowe,%Dane2%\Rodzaje_dok\Dokumentacja badań podłoża gruntowego.txt
FileAppend,DGI`nPodstawowe,%Dane2%\Rodzaje_dok\Dokumentacja Geologiczno Inżynierska.txt
FileAppend,DTR`nPodstawowe,%Dane2%\Rodzaje_dok\Karty katalogowe urządzeń i armatury.txt
FileAppend,DUS`nPodstawowe,%Dane2%\Rodzaje_dok\Raport o oddziaływaniu przedsięwzięcia na środowisko.txt
FileAppend,DW0`nPodstawowe,%Dane2%\Rodzaje_dok\Dokumentacja Wodno - Prawna.txt
FileAppend,HAZ`nPodstawowe,%Dane2%\Rodzaje_dok\HAZOP.txt
FileAppend,HYD`nPodstawowe,%Dane2%\Rodzaje_dok\Dokumentacja hydrologiczna.txt
FileAppend,IBP`nPodstawowe,%Dane2%\Rodzaje_dok\Instrukcja bezpieczeństwa pożarowego.txt
FileAppend,ID0`nPodstawowe,%Dane2%\Rodzaje_dok\Inwentaryzacja dendrologiczna.txt
FileAppend,INN`nPodstawowe,%Dane2%\Rodzaje_dok\Inna.txt
FileAppend,INS`nPodstawowe,%Dane2%\Rodzaje_dok\Instrukcja.txt
FileAppend,IPP`nPodstawowe,%Dane2%\Rodzaje_dok\Inwentaryzacja przyrodnicza.txt
FileAppend,KI0`nPodstawowe,%Dane2%\Rodzaje_dok\Kosztorys inwestorski.txt
FileAppend,KIP`nPodstawowe,%Dane2%\Rodzaje_dok\Karta informacyjna przedsięwzięcia.txt
FileAppend,KP0`nPodstawowe,%Dane2%\Rodzaje_dok\Koncepcja.txt
FileAppend,MPZ`nPodstawowe,%Dane2%\Rodzaje_dok\Miejscowy Plan Zagospodarowania Przestrzennego.txt
FileAppend,ODS`nPodstawowe,%Dane2%\Rodzaje_dok\Odstępstwa.txt
FileAppend,OE1`nPodstawowe,%Dane2%\Rodzaje_dok\Opracowanie Ewidencyjne.txt
FileAppend,OG1`nPodstawowe,%Dane2%\Rodzaje_dok\Opracowanie Geodezyjne.txt
FileAppend,OGE`nPodstawowe,%Dane2%\Rodzaje_dok\Opinia geotechniczna.txt
FileAppend,OK1`nPodstawowe,%Dane2%\Rodzaje_dok\Opracowanie Kartograficzne.txt
FileAppend,OWY`nPodstawowe,%Dane2%\Rodzaje_dok\Obliczenia wytrzymałościowe.txt
FileAppend,OZW`nPodstawowe,%Dane2%\Rodzaje_dok\Ocena zagrożenia wybuchem.txt
FileAppend,OWY`nPodstawowe,%Dane2%\Rodzaje_dok\Obliczenia wytrzymałościowe.txt
FileAppend,PB0`nPodstawowe,%Dane2%\Rodzaje_dok\Projekt Budowlany.txt
FileAppend,PGE`nPodstawowe,%Dane2%\Rodzaje_dok\Projekt Geotechczniczny.txt
FileAppend,PKB`nPodstawowe,%Dane2%\Rodzaje_dok\Plan kontroli i badań.txt
FileAppend,POR`nPodstawowe,%Dane2%\Rodzaje_dok\Projekt Organizacji Robót.txt
FileAppend,PPC`nPodstawowe,%Dane2%\Rodzaje_dok\Projekt prób ciśnieniowych.txt
FileAppend,PR0`nPodstawowe,%Dane2%\Rodzaje_dok\Przedmiar robót.txt
FileAppend,PRG`nPodstawowe,%Dane2%\Rodzaje_dok\Projekt robót geologicznych.txt
FileAppend,PTE`nPodstawowe,%Dane2%\Rodzaje_dok\Projekt technologiczny.txt
FileAppend,PW0`nPodstawowe,%Dane2%\Rodzaje_dok\Projekt Wykonawczy - brak etapowania.txt
FileAppend,PZA`nPodstawowe,%Dane2%\Rodzaje_dok\Plan zapobiegania awariom.txt
FileAppend,PZA`nPodstawowe,%Dane2%\Rodzaje_dok\Plan zapobiegania awariom.txt
FileAppend,RID`nPodstawowe,%Dane2%\Rodzaje_dok\Raport Inwentaryzacji Dendrologicznej.txt
FileAppend,RIT`nPodstawowe,%Dane2%\Rodzaje_dok\Raport Inwentaryzacji Terenowej.txt
FileAppend,ROB`nPodstawowe,%Dane2%\Rodzaje_dok\Raport o bezpieczeństwie.txt
FileAppend,ROS`nPodstawowe,%Dane2%\Rodzaje_dok\Raport oceny oddziaływaniu na środowisko.txt
FileAppend,STO`nPodstawowe,%Dane2%\Rodzaje_dok\Specyfikacja Techniczna.txt
FileAppend,SW0`nPodstawowe,%Dane2%\Rodzaje_dok\Specyfikacje Techniczne Wykonania i Odbioru Robót Budowlanych.txt
FileAppend,UPB`nPodstawowe,%Dane2%\Rodzaje_dok\Uzgodnienia do projektu budowlanego.txt
FileAppend,UZG`nPodstawowe,%Dane2%\Rodzaje_dok\Uzgodnienia.txt
FileAppend,WDL`nPodstawowe,%Dane2%\Rodzaje_dok\Wniosek o decyzję lokalizacyjną.txt
FileAppend,WN0`nPodstawowe,%Dane2%\Rodzaje_dok\Wykaz Nieruchomości.txt
FileAppend,WOG`nPodstawowe,%Dane2%\Rodzaje_dok\Wstępna opinia geotechniczna.txt
FileAppend,WOW`nPodstawowe,%Dane2%\Rodzaje_dok\Wnioski o wycinkę.txt
FileAppend,WS0`nPodstawowe,%Dane2%\Rodzaje_dok\Projekt Wstępny.txt
FileAppend,WTO`nPodstawowe,%Dane2%\Rodzaje_dok\Warunki Techniczne Wykonania i Odbioru.txt
FileAppend,ZK0`nPodstawowe,%Dane2%\Rodzaje_dok\Zbiorcze zestawienie kosztów.txt
FileAppend,ZUD`nPodstawowe,%Dane2%\Rodzaje_dok\Wniosek na naradę koordynacyjną.txt

FileAppend,01`nTechnologiczna,%Dane2%\licznik\01 - Schemat blokowy.txt
FileAppend,02`nTechnologiczna,%Dane2%\licznik\02 - Schemat PID.txt
FileAppend,01`nTechnologiczna,%Dane2%\licznik\01 - Schemat PFD.txt
FileAppend,01`nTechnologiczna,%Dane2%\licznik\01 - Plan sytuacyjny z LUW.txt
FileAppend,02`nTechnologiczna,%Dane2%\licznik\02 - Rysunki poglądowe.txt
FileAppend,03`nTechnologiczna,%Dane2%\licznik\03 - Zabudowa urządzenia.txt
FileAppend,04`nTechnologiczna,%Dane2%\licznik\04 - Rysunek założeniowy.txt
FileAppend,05`nTechnologiczna,%Dane2%\licznik\05 - Rysunki montażowe.txt
FileAppend,06`nTechnologiczna,%Dane2%\licznik\06 - Strefa zagrożenia wybuchem.txt
FileAppend,11`nTechnologiczna,%Dane2%\licznik\11 - Montaż lini (izometryk).txt
FileAppend,21`nTechnologiczna,%Dane2%\licznik\21 - Rysunki powtarzalne.txt
FileAppend,22`nTechnologiczna,%Dane2%\licznik\22 - Profil podłużny.txt
FileAppend,31`nTechnologiczna,%Dane2%\licznik\31 - Inne.txt
FileAppend,01`nLiniowa,%Dane2%\licznik\01 - Orientacja.txt
FileAppend,02`nLiniowa,%Dane2%\licznik\02 - PZT.txt
FileAppend,03`nLiniowa,%Dane2%\licznik\03 - Ortofotomapy.txt
FileAppend,04`nLiniowa,%Dane2%\licznik\04 - Profil.txt
FileAppend,05`nLiniowa,%Dane2%\licznik\05 - Skrzyżowania.txt
FileAppend,06`nLiniowa,%Dane2%\licznik\06 - Schematy.txt
FileAppend,07`nLiniowa,%Dane2%\licznik\07 - Inne.txt
FileAppend,00`nKonstrukcja,%Dane2%\licznik\00 - Rysunki informacyjne.txt
FileAppend,01`nKonstrukcja,%Dane2%\licznik\01 - Plany i rzuty.txt
FileAppend,02`nKonstrukcja,%Dane2%\licznik\02 - Rysunki kons. żelbetowych.txt
FileAppend,03`nKonstrukcja,%Dane2%\licznik\03 - Rysunki kons. stalowych.txt
FileAppend,04`nKonstrukcja,%Dane2%\licznik\04 - Rysunki kons. drewnianych.txt
FileAppend,05`nKonstrukcja,%Dane2%\licznik\05 - Rysunki pozostałe.txt
FileAppend,10`nKonstrukcja,%Dane2%\licznik\10 - Komory przewiertowe.txt
FileAppend,11`nKonstrukcja,%Dane2%\licznik\11 - Komory nadawcze.txt
FileAppend,12`nKonstrukcja,%Dane2%\licznik\12 - Komory odbiorcze.txt
FileAppend,20`nKonstrukcja,%Dane2%\licznik\20 - Zabezpieczenia liniowe.txt
FileAppend,21`nKonstrukcja,%Dane2%\licznik\21 - Zabezpieczenia ściankami szczelnymi.txt
FileAppend,22`nKonstrukcja,%Dane2%\licznik\22 - Zabezpieczenie ściankami systemowymi.txt
FileAppend,30`nKonstrukcja,%Dane2%\licznik\30 - Wzmocnienie gruntów słabonośnych.txt
FileAppend,40`nKonstrukcja,%Dane2%\licznik\40 - Rysunki pozostałe.txt

return

starter_wzor:

FileCreateDir, %Wzor2%\Model 3D
FileAppend,MOD,%Wzor2%\Model 3D\info.txt
FileCreateDir, %Wzor2%\Opis techniczny
FileAppend,DOC,%Wzor2%\Opis techniczny\info.txt
FileCreateDir, %Wzor2%\Schemat orurowania i oprzyrządowania
FileAppend,PID,%Wzor2%\Schemat orurowania i oprzyrządowania\info.txt
FileCreateDir, %Wzor2%\Projekt zagospodarowania terenu
FileAppend,PZT,%Wzor2%\Projekt zagospodarowania terenu\info.txt
FileCreateDir, %Wzor2%\Rysunek 2D
FileAppend,RYS,%Wzor2%\Rysunek 2D\info.txt
FileCreateDir, %Wzor2%\Załącznik
FileAppend,ZAL,%Wzor2%\Załącznik\info.txt
FileCreateDir, %Wzor2%\Schemat procesowy
FileAppend,PFD,%Wzor2%\Schemat procesowy\info.txt
FileCreateDir, %Wzor2%\Schemat
FileAppend,SCH,%Wzor2%\Schemat\info.txt

return

;_________________________________________________________Znaki specjalne___________________________________________

znaki_spec:
znaki_spec := "Użyto nie dozwolonych znaków: #%&{}\<>*?/$!':@+`|="
znak_test = 0
znaki_spec1 := "#"
IfInString, znaki_spec_spr, %znaki_spec1%
znak_test = 1
znaki_spec2 := "%"
IfInString, znaki_spec_spr, %znaki_spec2%
znak_test = 1
znaki_spec3 := "&"
IfInString, znaki_spec_spr, %znaki_spec3%
znak_test = 1
znaki_spec4 := "{"
IfInString, znaki_spec_spr, %znaki_spec4%
znak_test = 1
znaki_spec5 := "}"
IfInString, znaki_spec_spr, %znaki_spec5%
znak_test = 1
znaki_spec6 := "\"
IfInString, znaki_spec_spr, %znaki_spec6%
znak_test = 1
znaki_spec7 := "<"
IfInString, znaki_spec_spr, %znaki_spec7%
znak_test = 1
znaki_spec8 := ">"
IfInString, znaki_spec_spr, %znaki_spec8%
znak_test = 1
znaki_spec9 := "*"
IfInString, znaki_spec_spr, %znaki_spec9%
znak_test = 1
znaki_spec10 = "
IfInString, znaki_spec_spr, %znaki_spec10%
znak_test = 1	
znaki_spec11 := "?"
IfInString, znaki_spec_spr, %znaki_spec11%
znak_test = 1
znaki_spec12 := "/"
IfInString, znaki_spec_spr, %znaki_spec12%
znak_test = 1
znaki_spec13 := "$"
IfInString, znaki_spec_spr, %znaki_spec13%
znak_test = 1	
znaki_spec14 := "!"
IfInString, znaki_spec_spr, %znaki_spec14%
znak_test = 1
znaki_spec15 := "'"
IfInString, znaki_spec_spr, %znaki_spec15%
znak_test = 1
znaki_spec16 := "="
IfInString, znaki_spec_spr, %znaki_spec16%
znak_test = 1
znaki_spec17 := ":"
IfInString, znaki_spec_spr, %znaki_spec17%
znak_test = 1
znaki_spec18 := "@"
IfInString, znaki_spec_spr, %znaki_spec18%
znak_test = 1
znaki_spec19 := "+"
IfInString, znaki_spec_spr, %znaki_spec19%
znak_test = 1	
znaki_spec21 := "|"
IfInString, znaki_spec_spr, %znaki_spec21%
znak_test = 1
If (znak_test = 1)
	MsgBox, 16, Ostrzeżenie!, %znaki_spec%
return


;_________________________________________________________Funkcje___________________________________________

FindAndReplace( obj, search, replace )
{
	obj.Selection.Find.ClearFormatting
	obj.Selection.Find.Replacement.ClearFormatting
	obj.Selection.Find.Execute( search, 0, 0, 0, 0, 0, 1, 1, 0, replace, 2)
}