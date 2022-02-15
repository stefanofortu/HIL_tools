elenco = ["BH.BRAKE_FE3",
          "BH.ADASIS_V2",
          "BH.AIRBAG1",
          "BH.TSR_01",
          "BH.STATUS_NTP",
          "BH.CRUISE_FBK",
          "BH.BMS_LVB1",
          "BH.NIT_GPS_INFO",
          "BH.NFR_HMI",
          "BH.STATUS_DOORS",
          "BH.HAPTIC1",
          "BH.STATUS_C-LCU",
          "BH.TIME&DATE",
          "BH.TIME&DATE2",
          "BH.NIT_ADAS_SETUP2",
          "BH.NIT_ADAS_SETUP",
          "BH.STATUS_DASM",
          "BH.STATUS_DASM_2",
          "C1.C1.ASR0",
          "C1.NCR_INFO",
          "C1.YRS_DATA",
          "C1.LWS-STEERING_ANGLE_SENSOR",
          "C1.RWS_FEEDBACK",
          "C1.ASR3",
          "C1.NAM1",
          "C1.BRAKE_FE3",
          "C1.ADASIS_V2",
          "C1.NFR_RWS_TARGET",
          "C1.AIRBAG1",
          "C1.SH2_MR1",
          "C1.TSR_01",
          "C1.MOT1",
          "C1.MOT2",
          "C1.MOT2_V2",
          "C1.NVO_INFO",
          "C1.STATUS_NTP",
          "C1.BMS_LVB1",
          "C1.NIT_GPS_INFO",
          "C1.NFR_HMI",
          "C1.STATUS_DOORS",
          "C1.DVS_STATUS",
          "C1.ASR_PEDAL_PRESS",
          "C1.HAPTIC1",
          "C1.STATUS_B_CAN",
          "C1.STATUS_C-NCM",
          "C1.STATUS_C-NFR",
          "C1.STATUS_C-NCA_NCR",
          "C1.STATUS_C-NCM2",
          "C1.STATUS_B_CAN2",
          "C1.TIME&DATE2",
          "C1.TIME&DATE",
          "C1.EOL_CONFIGURATION_V4",
          "C1.NIT_ADAS_SETUP2",
          "C1.NIT_ADAS_SETUP",
          "C1.STATUS_DASM",
          "C1.STATUS_DASM_2",
          "C2.C1.ASR0",
          "C2.NCR_INFO",
          "C2.YRS_DATA",
          "C2.LWS-STEERING_ANGLE_SENSOR",
          "C2.RWS_FEEDBACK",
          "C2.BRAKE_FE3",
          "C2.ASR3",
          "C2.NAM1",
          "C2.NFR_RWS_TARGET",
          "C2.AIRBAG1",
          "C2.SH2_MR1",
          "C2.TSR_01",
          "C2.MOT1",
          "C2.MOT2",
          "C2.MOT2_V2",
          "C2.NVO_INFO",
          "C2.CRUISE_FBK",
          "C2.BMS_LVB1",
          "C2.NFR_HMI",
          "C2.STATUS_DOORS",
          "C2.DVS_STATUS",
          "C2.ASR_PEDAL_PRESS",
          "C2.STATUS_B_CAN",
          "C2.STATUS_C-NCM",
          "C2.STATUS_C-NFR",
          "C2.STATUS_C-NCA_NCR",
          "C2.STATUS_C-NCM2",
          "C2.STATUS_B_CAN2",
          "C2.STATUS_C-LCU",
          "C2.TIME&DATE2",
          "C2.TIME&DATE",
          "C2.EOL_CONFIGURATION_V4"]

elencoOutput = []
for msgName in elenco:
    duplicati = []
    duplicati.append(msgName)
    for msgNameDouble in elenco:
        if msgName[3:] == msgNameDouble[3:] and msgName[:3] != msgNameDouble[:3]:
            duplicati.append(msgNameDouble)
            # elencoOutput.append([msgName, msgNameDouble])
            elenco.remove(msgNameDouble)
    elencoOutput.append(duplicati)

for out in elencoOutput:
    # print(out)
    pass
# print(elencoOutput[2])
from itertools import combinations

diff_min = "" # str(10)
diff_max = "" # str(50)
nomeLog = "post_processing.blf"
can_name_present = True
for elem in elencoOutput:
    combinazioni = list(combinations(elem, 2))
    # print(combinazioni)
    for combo in combinazioni:
        # print("combo" + str(combo))
        msg1 = combo[0]
        msg2 = combo[1]
        outputList = []
        outputList.extend(["BLF[check_msg_presence]=" + nomeLog])
        # print(outputList)
        if can_name_present:
            outputList.append(msg1[:2])
        else:
            outputList.append("")
        outputList.append(msg1[3:])
        if can_name_present:
            outputList.append(msg2[:2])
        else:
            outputList.append("")
        outputList.append(msg2[3:])
        outputList.extend([diff_min, diff_max, ""])
        joined_string = ";".join(outputList)
        print(joined_string)

        # outputList.append(api)
        # outputList.append()
        # print("BLF[" + api + "]=" + nomeLog + ";")