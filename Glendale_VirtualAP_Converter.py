'''
Conversion Script (Meraki BSSID MAC Addresses):
LINK: https://documentation.meraki.com/MR/WiFi_Basics_and_Best_Practices/Calculating_Cisco_Meraki_BSSID_MAC_Addresses
-----------------
Models: MR12, 16, 18, 24, 62, 66
OUI:    00:18:OA
        88:15:44

Models: MR26, 32, 34, 74
OUI:    00:18:0A
        88:15:44
        E0:55:3D

Models: MR20, 30H, 33, 42, 42E, 52, 53, 53E, 70, 74, 84
OUI:    88:15:44
        E0:55:3D
        0C:8D:DB
        E0:CB:BC
'''

import xlsxwriter, csv
EXCEL_COUNTER = 0
SPACING_COUNTER = 0
MODEL_COUNTER = 0
SSID_COUNTER = 0
SSID_FORMAT = 1

mr_18 = ["0a", "1a"]
mr_26 = ["4a", "5a", "02"]
mr_32 = ["7d", "6d"]
mr_33 = ["bc", "ac"]
mr_66 = ["44", "54"]
mr_74 = ["db", "cb"]

list_2_4, list_5, names, mac_addresses, models = [], [], [], [], []

#load workbook
workbook = xlsxwriter.Workbook('Glendale_Network.xlsx')
worksheet = workbook.add_worksheet()


#Open Files with Names, Models, and MAC addresses of APs
with open('access points.csv', 'r') as f:
    csvReader = csv.reader(f, delimiter=',')
    accessPointdata = list(csvReader)

for i, row in enumerate(accessPointdata):

        names.append(accessPointdata[i][0])
        mac_addresses.append(accessPointdata[i][1])
        models.append(accessPointdata[i][2])

names.pop(0)
mac_addresses.pop(0)
models.pop(0)

def calc18(mac_address, name, mr):
    #Mac Address Breakdown
    mac_calc = mac_address[:2]
    mac_extractor_front = mac_address[2:6]
    mac_extractor_back = mac_address[8:]

    mr_18.insert(0, mac_extractor_front)
    mr_18.insert(1, mac_extractor_back)

    flag = True

    First_Pos_Hex = first_hex_calc(mac_calc, mr, flag)

    #Initial Mac Addresses
    list_2_4.append(mac_address)
    list_5.append(First_Pos_Hex + mr_18[0] + mr_18[3] + mr_18[1])

    flag = False
    #First Time Calculation hex updated
    hex_updater = first_hex_calc(mac_calc, mr, flag)

    list_2_4.append(hex_updater + mr_18[0] + mr_18[2] + mr_18[1])
    list_5.append(hex_updater + mr_18[0] + mr_18[3] + mr_18[1])

    #Update self, after appending mac addresses range(13)
    for i in range(14):
        counter = i + 1
        hex_updater = hex_calc(hex_updater, mr)
        list_2_4.append(hex_updater + mr_18[0] + mr_18[2] + mr_18[1])
        list_5.append(hex_updater + mr_18[0] + mr_18[3] + mr_18[1])

    for i in range(len(list_2_4)):
        counter = i + 1
        print(str(counter) + ") 2.4ghz: " + list_2_4[i] + "\t\t" + str(counter) + ") 5 Ghz: " + list_5[i])


    for i in range(len(list_2_4)):
        writer(name, mac_address, list_2_4[i], list_5[i])

    mr_18.pop(0)
    mr_18.pop(0)

    del list_2_4[:]
    del list_5[:]

def calc26(mac_address, name, mr):
    #Mac Address Breakdown
    mac_calc = mac_address[15:]
    mac_extractor_front = mac_address[2:6]
    mac_extractor_back = mac_address[8:15]

    mr_26.insert(0, mac_extractor_front)
    mr_26.insert(1, mac_extractor_back)

    #Initial Mac Addresses
    list_2_4.append(mr_26[4] + mr_26[0] + mr_26[2] + mr_26[1] + mac_calc)
    list_5.append(mr_26[4] + mr_26[0] + mr_26[3] + mr_26[1] + mac_calc)

    #Update self, after appending mac addresses range(14)
    for i in range(15):
        mac_calc = hex_calc(mac_calc, mr)
        list_2_4.append(mr_26[4] + mr_26[0] + mr_26[2] + mr_26[1] + mac_calc)
        list_5.append(mr_26[4] + mr_26[0] + mr_26[3] + mr_26[1] + mac_calc)

    for i in range(len(list_2_4)):
        counter = i + 1
        print(str(counter) + ") 2.4ghz: " + list_2_4[i] + "\t\t" + str(counter) + ") 5 Ghz: " + list_5[i])

    for i in range(len(list_2_4)):
        writer(name, mac_address, list_2_4[i], list_5[i])

    mr_26.pop(0)
    mr_26.pop(0)
    del list_2_4[:]
    del list_5[:]

def calc32(mac_address, name, mr):
    #Mac Address Breakdown
    mac_calc = mac_address[15:17]
    mac_extractor_front = "e2" + mac_address[2:6]
    mac_extractor_back = mac_address[8:15]

    mr_32.insert(0, mac_extractor_front)
    mr_32.insert(1, mac_extractor_back)


    #Initial Mac Addresses
    list_2_4.append(mr_32[0] + mr_32[2] + mr_32[1] + mac_calc)
    list_5.append(mr_32[0] + mr_32[3] + mr_32[1] + mac_calc)

    #Update self, after appending mac addresses range(14)
    for i in range(15):
        mac_calc = hex_calc(mac_calc, mr)
        list_2_4.append(mr_32[0] + mr_32[2] + mr_32[1] + mac_calc)
        list_5.append(mr_32[0] + mr_32[3] + mr_32[1] + mac_calc)

    for i in range(len(list_2_4)):
        counter = i + 1
        print(str(counter) + ") 2.4ghz: " + list_2_4[i] + "\t" + str(counter) + ") 5 Ghz: " + list_5[i] )

    for i in range(len(list_2_4)):
        writer(name, mac_address, list_2_4[i], list_5[i])
    mr_32.pop(0)
    mr_32.pop(0)
    del list_2_4[:]
    del list_5[:]

def calc33(mac_address, name, mr):
    #Mac Address Breakdown
    mac_calc = mac_address[:2]
    mac_extractor_front = mac_address[2:6]
    mac_extractor_back = mac_address[8:]
    counter = 1

    mr_33.insert(0, mac_extractor_front)
    mr_33.insert(1, mac_extractor_back)

    flag = True

    initial_5 = first_hex_calc(mac_calc, mr,flag)

    #Initial Mac Addresses
    list_2_4.append(mac_calc + mr_33[0] + mr_33[2] +  mr_33[1])
    list_5.append(initial_5 + mr_33[0] + mr_33[3] + mr_33[1])

    #range(14)
    for i in range(15):
        counter = 1 + i
        mac_calc = calc33_hex_calc(mac_calc, counter)
        list_2_4.append(mac_calc + mr_33[0] + mr_33[2] +  mr_33[1])
        list_5.append(mac_calc + mr_33[0] + mr_33[3] + mr_33[1])


    for i in range(len(list_2_4)):
        writer(name, mac_address, list_2_4[i], list_5[i])

    mr_33.pop(0)
    mr_33.pop(0)
    del list_2_4[:]
    del list_5[:]

def calc66(mac_address, name):
    #Mac Address Breakdown
    mac_calc = mac_address[:2]
    mac_extractor_front = mac_address[2:6]
    mac_extractor_back = mac_address[8:]

    #First 5Ghz MAC Setup (+0x02 Hexadecimal)
    First_5Ghz_66 = int(mac_calc, 16)
    First_5Ghz_66 = First_5Ghz_66 + 2
    First_5Ghz_66 = hex(First_5Ghz_66)
    First_5Ghz_66 = slice_hex(First_5Ghz_66)

    mr_66.insert(0, mac_extractor_front)
    mr_66.insert(1, mac_extractor_back)

    #Initial Mac Addresses
    list_2_4.append(mac_calc + mr_66[0] + mr_66[2] + mr_66[1])
    list_5.append(First_5Ghz_66 + mr_66[0] + mr_66[3] + mr_66[1])

    #Append all calculations range(14)
    for i in range(15):
        counter = 1 + i
        mac_calc = calc66_hex_calc(mac_calc, counter)
        list_2_4.append(mac_calc + mr_66[0] + mr_66[2] +  mr_66[1])
        list_5.append(mac_calc + mr_66[0] + mr_66[3] + mr_66[1])

    #Display all Mac Addresses after Calculations
    for i, index in enumerate(list_2_4):
        counter = 1+i
        print(str(counter) + ") 2.4ghz: " + list_2_4[i] + "\t" + str(counter) + ") 5 Ghz: " + list_5[i] )
        # writer(name, mac_address, list_2_4[i], list_5[i])

    for i in range(len(list_2_4)):
        writer(name, mac_address, list_2_4[i], list_5[i])

    mr_66.pop(0)
    mr_66.pop(0)
    del list_2_4[:]
    del list_5[:]

def calc74(mac_address, name):
    #Mac Address Breakdown
    mac_calc = mac_address[:2]  #0c
    mac_extractor_front = mac_address[2:6]  # :8d:
    mac_extractor_back = mac_address[8:]    #    #First 5Ghz MAC Setup (+0x02 Hexadecimal)

    First_5Ghz_74 = int(mac_calc, 16)
    First_5Ghz_74 = First_5Ghz_74 + 2
    First_5Ghz_74 = hex(First_5Ghz_74)
    First_5Ghz_74 = slice_hex(First_5Ghz_74)

    mr_74.insert(0, mac_extractor_front)
    mr_74.insert(1, mac_extractor_back)

    #Append First Special Case SSIDs
    list_2_4.append(mac_calc + mr_74[0] + mr_74[2] + mr_74[1])
    list_5.append(First_5Ghz_74 + mr_74[0] + mr_74[3] + mr_74[1])

    #Append all calculations range(14)
    for i in range(15):
        counter = 1 + i
        mac_calc = calc74_hex_calc(mac_calc, counter)
        list_2_4.append(mac_calc + mr_74[0] + mr_74[2] + mr_74[1])
        list_5.append(mac_calc + mr_74[0] + mr_74[3] + mr_74[1])

    #Display all Mac Addresses after Calculations
    for i, index in enumerate(list_2_4):
        counter = 1+i
        print(str(counter) + ") 2.4ghz: " + list_2_4[i] + "\t" + str(counter) + ") 5 Ghz: " + list_5[i] )
        # writer(name, mac_address, list_2_4[i], list_5[i])

    for i in range(len(list_2_4)):
        writer(name, mac_address, list_2_4[i], list_5[i])

    mr_74.pop(0)
    mr_74.pop(0)
    del list_2_4[:]
    del list_5[:]

#Create Text File and place results inside

def writer(name, physical_mac, entry_24, entry_5):

    global EXCEL_COUNTER #0
    global MODEL_COUNTER #0
    global SPACING_COUNTER #0
    global SSID_COUNTER #0
    global SSID_FORMAT #0

    #16th AP Name
    if (EXCEL_COUNTER % 17) == 0 and EXCEL_COUNTER != 0:
        worksheet.write(SPACING_COUNTER, 3, "MESH NETWORK MAC")
        SPACING_COUNTER += 1


    #Name, MAC, Model
    if (EXCEL_COUNTER % 17) == 0:
        worksheet.write(EXCEL_COUNTER, 0, name)
        worksheet.write(EXCEL_COUNTER, 1, physical_mac)
        worksheet.write(EXCEL_COUNTER, 2, models[MODEL_COUNTER])
        SSID_COUNTER += 1
        EXCEL_COUNTER += 1
        MODEL_COUNTER += 1


    #print virtual ssids in columns C = 2.4, D = 5

    worksheet.write(EXCEL_COUNTER, 4, entry_24)
    worksheet.write(EXCEL_COUNTER, 5, entry_5)
    worksheet.write(SSID_COUNTER, 3, "SSID " + str(SSID_FORMAT) + ": ")

    if SSID_FORMAT == 1:
        worksheet.write(SSID_COUNTER, 3, "GLAC_PATRON")
    if SSID_FORMAT == 2:
        worksheet.write(SSID_COUNTER, 3, "FIRING_RANGE")
    if SSID_FORMAT == 3:
        worksheet.write(SSID_COUNTER, 3, "COG_STAFF")
    if SSID_FORMAT == 4:
        worksheet.write(SSID_COUNTER, 3, "COG_GUEST")
    if SSID_FORMAT == 5:
        worksheet.write(SSID_COUNTER, 3, "GFD_GUEST")
    if SSID_FORMAT == 6:
        worksheet.write(SSID_COUNTER, 3, "Treasurer")
    if SSID_FORMAT == 7:
        worksheet.write(SSID_COUNTER, 3, "COG_HVAC")
    if SSID_FORMAT == 8:
        worksheet.write(SSID_COUNTER, 3, "COG_BYOD")
    if SSID_FORMAT == 9:
        worksheet.write(SSID_COUNTER, 3, "COG_DEVICES")
    if SSID_FORMAT == 10:
        worksheet.write(SSID_COUNTER, 3, "VJC_GUEST")
    if SSID_FORMAT == 11:
        worksheet.write(SSID_COUNTER, 3, "J135_GUEST")
    if SSID_FORMAT == 12:
        worksheet.write(SSID_COUNTER, 3, "Sparr")
    if SSID_FORMAT == 13:
        worksheet.write(SSID_COUNTER, 3, "GPD_GUEST")
    if SSID_FORMAT == 14:
        worksheet.write(SSID_COUNTER, 3, "MedixSafe")
    if SSID_FORMAT == 15:
        worksheet.write(SSID_COUNTER, 3, "GUEST")

    SSID_FORMAT += 1
    EXCEL_COUNTER += 1
    SPACING_COUNTER += 1
    SSID_COUNTER += 1

    if SSID_FORMAT == 16:
        SSID_FORMAT = 0

def calc33_hex_calc(hexi, counter):
    num = int(hexi, 16)

    if counter == 1:
        final = num + 6
        num = hex(final)
        return slice_hex(num)
    if counter == 2 or counter <= 7:
        final = num + 4
        num = hex(final)
        return slice_hex(num)
    if counter == 8 or counter == 15:
        final = num - 60
        num = hex(final)
        return slice_hex(num)
    if counter >= 9 or counter <= 14:
        final = num + 4
        num = hex(final)
        return slice_hex(num)
    #16th rotation
    # if counter == 15:
    #     final = num - 60
    #     num = hex(final)
    #     return slice_hex(num)

def calc66_hex_calc(hexi, counter):
    num = int(hexi, 16)
    if counter == 1:
        final = num + 6
        num = hex(final)
        return slice_hex(num)
    if counter == 2 or counter == 6 or counter == 10 or counter == 14:
        final = num - 12
        num = hex(final)
        return slice_hex(num)
    if counter == 3 or counter == 5 or counter == 7 or counter == 9 or counter == 11 or counter == 13 or counter == 15:
        final = num + 4
        num = hex(final)
        return slice_hex(num)
    if counter == 4 or counter == 8 or counter == 12:
        final = num + 20
        num = hex(final)
        return slice_hex(num)
    # if counter == 5:
    #     final = num + 4
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 6:
    #     final = num - 12
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 7:
    #     final = num + 4
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 8:
    #     final = num + 20
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 9:
    #     final = num + 4
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 10:
    #     final = num - 12
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 11:
    #     final = num + 4
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 12:
    #     final = num + 20
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 13:
    #     final = num + 4
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 14:
    #     final = num - 12
    #     num = hex(final)
    #     return slice_hex(num)
    #16th rotation
    # if counter == 15:
    #     final = num + 4
    #     num = hex(final)
    #     return slice_hex(num)

def calc74_hex_calc(hexi, counter):
    num = int(hexi, 16)

    if counter == 1:
        final = num - 2
        num = hex(final)
        return slice_hex(num)
    if counter == 2 or counter == 3 or counter == 5 or counter == 6 or counter == 7 or counter == 9 or counter == 10 or counter == 11 or counter == 13 or counter == 14 or counter == 15:
        final = num - 4
        num = hex(final)
        return slice_hex(num)
    if counter == 4 or counter == 8 or counter == 12:
        final = num + 28
        num = hex(final)
        return slice_hex(num)
    # if counter == 5 or counter == 6 or counter == 7:
    #     final = num - 4
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 8:
    #     final = num + 28
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 9 or counter == 10 or counter == 11:
    #     final = num - 4
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 12:
    #     final = num + 28
    #     num = hex(final)
    #     return slice_hex(num)
    # if counter == 13 or counter == 14:
    #     final = num - 4
    #     num = hex(final)
    #     return slice_hex(num)
    #16th rotation
    # if counter == 15:
    #     final = num - 4
    #     num = hex(final)
    #     return slice_hex(num)

def first_hex_calc(hexi, mr, flag):
    #converted into number
    num = int(hexi, 16)

    if mr == "MR33":
        first = num + 2
        num = hex(first)
        final_conversion = slice_hex(num)
        return final_conversion

    #if mr = 18 add 6 to first octet
    if mr == "MR18":

        if flag == True:
        	first = num + 2
        	num = hex(first)
        	final_conversion = slice_hex(num)
        	return final_conversion
        final = num + 6
        num = hex(final)
        final_conversion = slice_hex(num)
        return final_conversion

def hex_calc(hexi, mr):
	#converted into number
	num = int(hexi, 16)

	#if mr = 18 add 4 to first octet
	if mr == "MR18":
		final = num + 4
		num = hex(final)
		return slice_hex(num)

	#if mr = 32 add 1 to 6th octet
	if mr == "MR32" or mr == "MR26":
	    final = num + 1
	    num = hex(final)
	    return slice_hex(num)

def slice_hex(r_hex):
    result = r_hex[2:]
    if len(result) != 2:
        result = "0" + result

    return result

for num, mod in enumerate(models):
	if mod == "MR18":
		calc18(mac_addresses[num], names[num], mod)
	if mod == "MR26":
		calc26(mac_addresses[num], names[num], mod)
	if mod == "MR32":
		calc32(mac_addresses[num], names[num], mod)
	if mod == "MR33":
		calc33(mac_addresses[num], names[num], mod)
	if mod == "MR66":
		calc66(mac_addresses[num], names[num])
	if mod == "MR74":
		calc74(mac_addresses[num], names[num])

workbook.close()
