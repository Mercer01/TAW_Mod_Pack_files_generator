import os, time

import xlwt
from openpyxl import *

wb = xlwt.Workbook()

last_checked_date = ""
fileslist = []
mod_pack_sheet = wb.add_sheet('Modpack', cell_overwrite_ok=True)
mod_pack_sheet.write(0, 0, "AM2_Normal_Modpack_Summary")
mod_pack_sheet.write(3, 1, "ModpackName")
file_counter = 0
modfileandsize = []
files_sheet = wb.add_sheet('Files', cell_overwrite_ok=True)
modnames = []
mods_index = [
				("ace", "Advanced Combat Realism", "3.9.8", ""),
				("MG8", "Might Gau 8 Gun Mod", "", ""),
				("cha_av8b", "Harrier Jump jet/AVB", ""),
				("EAWS_EF2000", "EuroFigther Typhoon", ""),
				("FIR_A10", "A10A + A10C", ""),
				("FIR_F14", "F-14 Tomcat", ""),
				(("FIR_AWS", "FIR_AirWeaponSystem"), "FIR Compatability", ""),
				("globemaster_c17", "C17 Globemaster", ""),
				("js_jc_fa18", "F/A - 18", ""),
				(("MELB", "taw_melb_compat"), "MELB Littlebirds", ""),
				("po_", "Project Opfor", ""),
				("rhsgref", "Red Hammer Studios: GREF", ""),
				("rhsusf", "Red Hammer Studios: USAF", ""),
				("rhssaf", "Red Hammer Studios: SAF", ""),
				("rhs_", "Red Hammer Studios: AFRF", ""),
				(("UK3CB_BAF_Vehicles", "uk3cb_baf"), "BAF: Vehicles", ""),
				(("sma", "SMA"), "SMA Weapons Pack", ""),
				("wop_gui", "Wop_GUI"), #Find Mod for this file
				(("AOR1", "aor2", "aorU", "MCB", "OPXT_TFAR"), "Uniforms", ""),
				(("bamse_cTab_fix", "cTab"), "cTAB - Additional Tablets", ""),
				("FIR_PilotCrewPack_US", "Pilot Uniforms", ""),
				("hlc", "N1Arms Guns And Compatability", ""),
				("insignia_addon", "TAW Insignias", ""),
				("rksl", "RKSL Attachments Pack", ""),
				("SC_SCAR", "Mk16 + Mk17 SCAR", ""),
				("taw_squad_box", "TAW Squad Box", ""),
				("VSM", "VSM Uniform Additions", ""),
				("zade_boc", "BackPackOnChestMod", ""),
				("cup_terrains", "CUP Terrains", ""),
				("AR", "Advanced Rappelling", ""),
				("cba", "CBA", ""),
				("SA_AdvancedSlingLoading", "Advanced Sling Loading", ""),
				("SA_AdvancedTowing", "Advanced Towing", ""),
				("stui", "Shack Tactial UI", ""),
				("task_force_radio", "Task Force ArrowHead Radio", ""),
				("viewDistance_TAW", "viewDistanceScript", ""),
				(("logo", "mod.cpp", "bikey"), "Other files")
]

for root, dirs, files in os.walk('E:\Desktop\Taw_modpack'):
	for name in files:
		path = root + "\\"+ name
		#print(name, os.path.getsize(root + "/"+ name), os.path.getmtime(path),name.split(".")[1])
		fileslist.append((name, os.path.getsize(root + "/"+ name), os.path.getmtime(path), os.path.splitext(path)[1], path))


for x, char in enumerate(fileslist):
	modfileandsize.append((char[0], char[1]))
	for count, value  in enumerate(char):
		#print(value)
		files_sheet.write(x,count+1, value)
		if "@" in str(value):
			#print((x,0,value.split('\\')))
			for character in value.split('\\'):
				if "@" in str(character):
					files_sheet.write(x, 0 ,character)
					if character not in modnames:
						modnames.append(character)
#Write Summary

#Write Content and Modpacks

for value, modname in enumerate(modnames):
	mod_pack_sheet.write(3+value, 0, modname)

mod_pack_sheet.write(2, 0, "Mod Name")
mod_pack_sheet.write(2, 1, "Total Mod Size (GB)")
mod_pack_sheet.write(2, 3, "Date Checked")
mod_pack_sheet.write(2, 4, time.strftime("%d/%m/%Y"))

mod_pack_sheet.write(3, 1, "=SUMIFS(Files!C:C,Files!A:A,Modpack!A4)")
mod_pack_sheet.write(4, 1, "=SUMIFS(Files!C:C,Files!A:A,Modpack!A5)")
mod_pack_sheet.write(5, 1, "=SUMIFS(Files!C:C,Files!A:A,Modpack!A6)")
mod_pack_sheet.write(6, 1, "=SUMIFS(Files!C:C,Files!A:A,Modpack!A7")
mod_pack_sheet.write(7, 0, "Total")
mod_pack_sheet.write(7, 1, "=SUM(B4:B7)/1024/1024/1024")

mod_pack_sheet.write(1, 5, "Mod name")
mod_pack_sheet.write(1, 6, "Mod Size (MB)")
mod_pack_sheet.write(1, 7, "Version")
mod_pack_sheet.write(1, 8, "Requires")

mod_pack_sheet.write(2, 5, "Total")
mod_pack_sheet.write(2, 6, "=SUM(G5:G39)")



modandfilesize = []

for modname in mods_index:
	modsize = 0
	if type(modname[0]) is str:
		#handle as string as only one file extension
		for filenameandsize in modfileandsize:
			modfiletocheck = modname[0]
			if len(modfiletocheck) <= len(filenameandsize[0]):
				if modfiletocheck == filenameandsize[0][:len(modfiletocheck)]:
					modsize += int(filenameandsize[1])
					#print(filenameandsize[0],modfiletocheck, )
	else:
		#print(modname)
		for mod in modname[0]:
			for filenameandsize in modfileandsize:
				modfiletocheck = mod
				if len(modfiletocheck) <= len(filenameandsize[0]):
					if modfiletocheck == filenameandsize[0][:len(modfiletocheck)]:
						modsize += int(filenameandsize[1])
						#print(filenameandsize[0],modfiletocheck)
					
			#handle as list as multiple file extensiosn
	print(modname[1], modsize/1024/1024, modname[2])
	modandfilesize.append((modname[1], modsize/1024/1024, modname[2]))
	

for count,x in enumerate(modandfilesize):
	mod_pack_sheet.write(4+count,5,x[0])
	mod_pack_sheet.write(4+count,6,x[2])
	if x[1] != 0:
		mod_pack_sheet.write(4+count, 7, x[1])

wb.save('TAW_AM2_Modpack_sheet.xls')
