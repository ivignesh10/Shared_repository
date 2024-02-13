import xml.etree.ElementTree as ET
import logging
import os
import datetime
import argparse
import pandas as pd
import PySimpleGUI as sg

# Parse the xml file
def parse_xml(xml_file,outfolder):
    try:
        tree = ET.parse(xml_file)
        logging.info("XML is validated and its vaild")
        print("XML is validated and its vaild")
        root = tree.getroot()
        try:
            fetch_values(root,outfolder)
        except Exception as e:
            logging.error(f"Error in fetch values: {e}")
            sg.popup(f"The choosen XML file is invalid: {e}") if GUI else None
    except Exception as e:
        logging.error(f"The choosen XML file is invalid: {e}")
        sg.popup(f"The choosen XML file is invalid: {e}") if GUI else None

# Fetch values from xml and store in xlsx file
def fetch_values(root,outfolder):
    container = []
    ecuc_container_values = root.findall(".//{http://autosar.org/schema/r4.0}CONTAINERS/{http://autosar.org/schema/r4.0}ECUC-CONTAINER-VALUE")
    logging.info("Extracting data of container's and sub-container's SHORT NAME, DEFINATION-REF in progress")
    print("Extracting data of container's and sub-container's SHORT NAME, DEFINATION-REF in progress")
    for ecuc_container_value in ecuc_container_values:
        short_name = ecuc_container_value.find("./{http://autosar.org/schema/r4.0}SHORT-NAME")
        definition = ecuc_container_value.find("./{http://autosar.org/schema/r4.0}DEFINITION-REF")
        sub_containers = ecuc_container_value.findall("./{http://autosar.org/schema/r4.0}SUB-CONTAINERS/{http://autosar.org/schema/r4.0}ECUC-CONTAINER-VALUE")
        for sub_container in sub_containers:
            sub_shortname = sub_container.find("./{http://autosar.org/schema/r4.0}SHORT-NAME")
            sub_definition = sub_container.find("./{http://autosar.org/schema/r4.0}DEFINITION-REF")
            container.append({"Container short name":short_name.text, "Container definition": definition.text, "Subcontainer short name":sub_shortname.text, "Subcontainer definition":sub_definition.text})
    logging.info("Extraction is done")
    print("Extraction is done")
    df = pd.DataFrame(container)
    df.to_excel(f"{outfolder}/Autosar_output.xlsx", sheet_name = "KPIT")
    logging.info(f"Data saved in {outfolder}/Autosar_output.xlsx")
    print(f"Data saved in {outfolder}/Autosar_output.xlsx")
    sg.popup(f"Done! The output file is saved as {outfolder}/Autosar_output.xlsx") if GUI else None

# Configuring basic logging
logging.basicConfig(filename = f"KPIT_xml_parser_logfile_{datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S_%f")}.log", format = '%(asctime)s - %(message)s', level = logging.DEBUG)

if __name__ == "__main__":
    #Command line argurments halding
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--file", type = str, help = "Input xml files path")
    parser.add_argument("-o", "--output_folder", type = str, help = "Output directory path to save a xlsx file")
    arg = parser.parse_args()
    
    # Check the tool is executed through CLI with input arguments
    if arg.file and arg.output_folder:
        # Check if the paths are valid
        if os.path.isfile(arg.file) and os.path.isdir(arg.output_folder):
            GUI = False
            parse_xml(arg.file,arg.output_folder)
        else:
            logging.error("Error! Please provide a valid input file and output folder.")
            print("Error! Please provide a valid input file and output folder.")
  
    #if it is not executed through CLI lanch the GUI and get the user inputs
    else:
        # Define the layout of the GUI
        layout = [
            [sg.Text("Select the input Autosar xml file:", expand_x=True)],
            [sg.Input(key="-INFILE-", expand_x=True), sg.FileBrowse(file_types=(("XML Files", "*.xml"),))],
            [sg.Text("Select the output folder:", expand_x=True)],
            [sg.Input(key="-OUTFOLDER-", expand_x=True), sg.FolderBrowse()],
            [sg.Button("Start", expand_x=True), sg.Button("Exit", expand_x=True)],
            [sg.Output(size=(40, 10),key="output", expand_x=True, expand_y=True)]
        ]
        # Create the window object
        window = sg.Window("KPIT Autosar xml parser", layout, resizable=True)
        # Event loop to process user input
        while True:
            event, values = window.read()
            # If user clicks Exit or closes the window, end the program
            if event == "Exit" or event == sg.WIN_CLOSED:
                break
            # If user clicks Start, process the input file and write to the output file
            elif event == "Start":
                # Get the input file path and output folder path from the values dictionary
                infile = values["-INFILE-"]
                outfolder = values["-OUTFOLDER-"]
                # Check if the paths are valid
                if os.path.isfile(infile) and os.path.isdir(outfolder):
                    GUI = True
                    parse_xml(infile,outfolder)
                # If the paths are not valid, show a message box to warn the user
                else:
                    sg.popup("Error! Please select a valid input file and output folder.")
        # Close the window
        window.close()