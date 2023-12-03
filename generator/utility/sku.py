import logging


def standardize_sku(df):
    logging.debug("Standardize SKU")

    # Kabel Jumper
    df['SKU'] = df['SKU'].str.replace('ITBISA-JUMPER-MM-10', 'ITBISA-JUMPER-MM-10CM')
    df['SKU'] = df['SKU'].str.replace('ITBISA-JUMPER-MF-10', 'ITBISA-JUMPER-MF-10CM')
    df['SKU'] = df['SKU'].str.replace('ITBISA-JUMPER-FF-10', 'ITBISA-JUMPER-FF-10CM')
    df['SKU'] = df['SKU'].str.replace('CMCM', 'CM')

    # Arduino UNO
    df['SKU'] = df['SKU'].str.replace('ITBISA-ARDUINO-UNO-R3-CLONE', 'ITBISA-ARDUINO-UNO-R3-328P-SMD')
    df['SKU'] = df['SKU'].str.replace('ITBISA-ARDUINO-UNO-R3-CLONE-ONLY', 'ITBISA-ARDUINO-UNO-R3-328P-SMD')
    df['SKU'] = df['SKU'].str.replace('ITBISA-ARDUINO-UNO-R3-328P-SMD-ONLY', 'ITBISA-ARDUINO-UNO-R3-328P-SMD')
    df['SKU'] = df['SKU'].str.replace('ITBISA-ARDUINO-USB-A-TO-B-50CM', 'ITBISA-USB-A-TO-B-50CM')
    df['SKU'] = df['SKU'].str.replace('ITBISA-ARDUINO-USB-A-TO-B-30CM', 'ITBISA-USB-A-TO-B-30CM')
    df['SKU'] = df['SKU'].str.replace('ITBISA-USB-A-TO-B-ARDUINO-50CM', 'ITBISA-USB-A-TO-B-50CM')
    df['SKU'] = df['SKU'].str.replace('ITBISA-USB-A-TO-B-ARDUINO-30CM', 'ITBISA-USB-A-TO-B-30CM')
    df['SKU'] = df['SKU'].str.replace('-ONLY-ONLY', '-ONLY')

    # Transistor 2N2222
    df['SKU'] = df['SKU'].str.replace('ITBISA-NPN-2N2222A', 'ITBISA-NPN-2N2222A-TO92')
    df['SKU'] = df['SKU'].str.replace('-TO92-TO92', '-TO92')

    # Socket IC
    df['SKU'] = df['SKU'].str.replace('ITBISA-SOCKET-IC-DIP8-2X4-ROUNDHOLE', 'ITBISA-SOCKET-IC-DIP8-ROUNDHOLE')

    # Handphone
    df['SKU'] = df['SKU'].str.replace('ITBISA-XIAOMI-REDMI-NOTE-5A-TAM-NEW', 'ITBISA-XIAOMI-REDMI-NOTE-5A-TAM')

    # Senter
    df['SKU'] = df['SKU'].str.replace('ITBISA-SENTER-KUNING-TIGERHEAD-FT300', 'ITBISA-TIGERHEAD-SENTER-KUNING-FT300')
    df['SKU'] = df['SKU'].str.replace('ITBISA-SENTER-KUNING-TIGERHEAD-FT300', 'ITBISA-TIGERHEAD-SENTER-KUNING-FT300')

    # Bohlam
    df['SKU'] = df['SKU'].str.replace('ITBISA-RAPID-BULB-E10-2.5V-SCREW', 'ITBISA-RAPID-BOHLAM-E10-2.5V-SCREW')

    # Baterai
    df['SKU'] = df['SKU'].str.replace('4808818178', 'ITBISA-ABC-BATERAI-D-POWER')

    # 7 Segment
    df['SKU'] = df['SKU'].str.replace('ITBISA-7SEGMENT-CA-RED-0.56-1BIT', 'ITBISA-7SEGMENT-ANODE-RED-0.56-1BIT')
    df['SKU'] = df['SKU'].str.replace('ITBISA-7SEGMENT-CC-RED-0.56-1BIT', 'ITBISA-7SEGMENT-CATHODE-RED-0.56-1BIT')
    df['SKU'] = df['SKU'].str.replace('ITBISA-7SEGMENT-ANODE-0.56-1BIT', 'ITBISA-7SEGMENT-ANODE-RED-0.56-1BIT')
    df['SKU'] = df['SKU'].str.replace('ITBISA-7SEGMENT-CATHODE-0.56-1BIT', 'ITBISA-7SEGMENT-CATHODE-RED-0.56-1BIT')
