# eCTD Output File validator 
# Checks if all files that should be present by the index.xml and regional.xml are actually present in the output folder
# Written by Nikola Siladi, August, 2022.

# libraries used
from genericpath import isfile
import xlsxwriter as X1
import xlsxwriter as X2
import xlsxwriter as X3
import xlsxwriter as X4
from lxml import etree as ET1
import os
from pathlib import Path, PureWindowsPath

# reference for skipping backups
allowed_regions = []

# defining dictionary
dic = {}
file_dic = {}
# defining the Split function
def split_path(path_arg):
    path_array = path_arg.split('\\')
    return path_array

# assemble path of a file from file output
def assemble_main_path(definition_array, file_path):
    
    if file_path.startswith("..") == True:
        end_file_path = assemble_replace_path(file_path, True)
    else:
        end_file_path = file_path
    
    starts_with_sequence = False
    path_check = split_path(end_file_path)
    if path_check[0].startswith("0") == True:
        starts_with_sequence = True

    end_path = []
    end_path.append(definition_array[-3])
    end_path.append(definition_array[-2])
    if starts_with_sequence == False:
        end_path.append(definition_array[-1])
    end_path.append(end_file_path)

    return "/".join(end_path)

# popping the '..' part away from the replace path from index.xml
def assemble_replace_path(replace_path, is_regional = False):
    replace_array = replace_path.split('/')

    while replace_array[0].startswith(".."):
        replace_array.pop(0)

    return "/".join(replace_array)

# assembling the paths of files defined in regional.xml and index.xml
def assemble_regional_path(definition_array, full_path, file_path):

    if file_path.startswith(".."):
        end_file_path = assemble_replace_path(file_path, True)
    else:
        end_file_path = file_path

    
    starts_with_sequence = False
    path_check = split_path(end_file_path)
    if path_check[0].startswith("0") == True:
        starts_with_sequence = True
    

    end_path = []
    end_path.append(definition_array[-3])
    end_path.append(definition_array[-2])
    if starts_with_sequence == False:
        end_path.append(definition_array[-1])
    end_path.append(full_path)
    end_path.append(end_file_path)

    return "/".join(end_path)

def assemble_file_path(file_path):
    definition_array = split_path(file_path)
    while definition_array[0] not in allowed_regions:
        definition_array.pop(0)

    return "/".join(definition_array)


# defining the index path in an string
def return_index_path(path_arg):
    path_array = path_arg.split('\\')
    index_array = []
    index_array.append(path_array[-3])
    index_array.append(path_array[-2])

    return "/".join(index_array)

# defining the index path in an string
def return_dos_key(path_arg, isFileOutput = False):
    path_array = []

    # trimming the path to setting the region name as the root folder
    if isFileOutput:
        path_array = path_arg.split('\\')

        root_path_element = path_array[0]
        while root_path_element not in allowed_regions:
            root_path_element = path_array.pop(0)

    else:
        path_array = path_arg.split('/')

    # checkup to drop backups, validation reports, working documents sections, that aren't affected by the deletion bug

    # checkup to drop backups which are indicated by usually the wrong path

    if isFileOutput == False:
        check_val = "false"

        for region in allowed_regions:
            for path_segment in path_array:
                if region == path_segment:
                    check_val = "true"
                    break
        
        if check_val == "false":
            return "false"

    if isFileOutput:
        try:
            i = int(path_array[1])
        except ValueError:
            return "false"
        
    else:
        try:
            i = int(path_array[2])
        except ValueError:
            return "false"

    index_array = []
    if isFileOutput:
        index_array.append(path_array[0])
        index_array.append(path_array[1])
    else:
        index_array.append(path_array[1])
        index_array.append(path_array[2])

    output_string = "-".join(index_array)
    return output_string

# defining the write_excel_row function
def write_excel_row(worksheet, row, region, dossier_identifier, seq_no, file_path, file_id, operation):
    worksheet.write(row, 0, region)
    worksheet.write(row, 1, dossier_identifier)
    worksheet.write(row, 2, seq_no)
    worksheet.write(row, 3, file_path)
    worksheet.write(row, 4, file_id)
    worksheet.write(row, 5, operation)

# defining the insert into the dictionary

def write_dic_row(region, dossier_identifier, seq_no, file_path, file_id, operation):
    dos_key = return_dos_key(file_path)
    if dos_key != "false":
        if dos_key not in dic:
            dic[dos_key] = {
                'Region': [],
                'Dossier_ID': [],
                'Seq_No': [],
                'FilePath': [],
                'FileID': [],
                'Operation': []
            }

        dic[dos_key]['Region'].append(region)
        dic[dos_key]['Dossier_ID'].append(dossier_identifier)
        dic[dos_key]['Seq_No'].append(seq_no)
        dic[dos_key]['FilePath'].append(file_path)
        dic[dos_key]['FileID'].append(file_id)
        dic[dos_key]['Operation'].append(operation)

# defining the delete operator into the dictionary
def remove_dic_row(file_id):
    for dos_key in dic:
        for i, file_id_loop in enumerate(dic[dos_key]['FileID']):
            if file_id == dic[dos_key]['FileID'][i]:
                dic[dos_key]['Region'].pop(i)
                dic[dos_key]['Dossier_ID'].pop(i)
                dic[dos_key]['Seq_No'].pop(i)
                dic[dos_key]['FilePath'].pop(i)
                dic[dos_key]['FileID'].pop(i)
                dic[dos_key]['Operation'].pop(i)

# defining unique sequence key
def return_sequence_key(path_arg):
    path_array = path_arg.split('/')

    # checkup to drop backups which are indicated by usually the wrong path
    if path_array[0] not in allowed_regions:
        return "false"
    
    # checkup to drop backups, validation reports, working documents sections, that aren't affected by the deletion bug
    try:
        i = int(path_array[2])
    except ValueError:
        return "false"

    index_array = []
    index_array.append(path_array[0])
    index_array.append(path_array[1])
    index_array.append(path_array[2])

    output_string = "-".join(index_array)
    return output_string


# pulling out the file names from .xmls present in the inputted directory:
def extract_paths_from_xml(ectd_path):

    # defining Excel workbook and setting headers
    workbook = X1.Workbook('eCTD_XML_Report.xlsx')
    worksheet = workbook.add_worksheet("eCTD_Report")
    headers = ["Region", "Dossier Identifier", "Sequence number", "File path", "File ID", "Operation"]
    header_format = workbook.add_format({'bold': True})
    for i, header in enumerate(headers):
        worksheet.write(0, 0 + i, header, header_format)
    worksheet.freeze_panes(1, 0)
    
    row = 1

    file_count = 0
    del_file_count = 0

    sequences = []
    seq_cnt = 0

    # looping through the xmls in folders/subfolders(in this case, eCTD output folder xmls(both regional and main index.xmls))

    #replace the ectd_path string with your path of the extracted eCTDs, with folders

    pathlist = Path(ectd_path).glob('**/*.xml')

    for path in pathlist:
        pp = PureWindowsPath(path)

        # exclude all backup folders of any kind
        if 'ackup' in str(pp):
            continue

        # parsing XML file
        tree = ET1.parse(pp)
        root = tree.getroot()

        if pp.match('*regional.xml') == True:
            # first parent folder of regional.xml is country name. second one is m1. the third parent carries the sequence number folder path
            parent_dir = pp.parents[2]
            p_array = split_path(str(parent_dir))
            
            # Skip backups that weren't created by the system, but copy/pasted by the user
            if p_array[-3] not in allowed_regions:
                continue 
            
            # if sequence is not composed only of numbers(backups and user-created/backup folders)
            try:
                i = int(p_array[-1])
            except ValueError:
                continue
            
            # manipulating regional XML

            # new operator
            for file_new in root.findall('.//leaf[@operation="new"]'):
                file_path = file_new.attrib['href']
                file_id = file_new.attrib['ID']

                full_path = return_index_path(str(path))

                end_path_str = assemble_regional_path(p_array, full_path, file_path)

                if end_path_str.endswith(".xml") == True:
                    continue

                write_dic_row(p_array[-3], p_array[-2], p_array[-1], end_path_str, file_id, "New")
                file_count += 1
            
            # replace operator
            for file_replace in root.findall('.//leaf[@operation="replace"]'):
                file_path = file_replace.attrib['href']
                file_id = file_replace.attrib['ID']

                full_path = return_index_path(str(path))

                end_path_str = assemble_regional_path(p_array, full_path, file_path)

                if end_path_str.endswith(".xml") == True:
                    continue
                
                write_dic_row(p_array[-3], p_array[-2], p_array[-1], end_path_str, file_id, "Replace")

                file_count += 1
            
            # delete operator
            for file_delete in root.findall('.//leaf[@operation="delete"]'):
                file_id = file_delete.attrib['ID']
            
                remove_dic_row(file_id)

                file_count -= 1
                del_file_count += 1

        # else it's index.xml with parent folder named as the sequence number
        else:
            parent_dir = pp.parents[0]
            p_array = split_path(str(parent_dir))

            # Skip backups that weren't created by the system, but copy/pasted by the user
            if p_array[-3] not in allowed_regions:
                continue

            # if sequence is not composed only of numbers(backups and user-created/backup folders)
            try:
                i = int(p_array[-1])
            except ValueError:
                continue

            # manipulating main XML 

            # new operator
            for file_new in root.findall('.//leaf[@operation="new"]'):
                file_path = file_new.attrib['href']
                file_id = file_new.attrib['ID']

                end_path_str = assemble_main_path(p_array, file_path)

                if file_path.endswith(".xml") == True:
                    continue
                
                write_dic_row(p_array[-3], p_array[-2], p_array[-1], end_path_str, file_id, "New")
                file_count += 1
            
            # replace operator
            for file_replace in root.findall('.//leaf[@operation="replace"]'):
                file_path = file_replace.attrib['href']
                file_id = file_replace.attrib['ID']

                end_path_str = assemble_main_path(p_array, file_path)

                if file_path.endswith(".xml") == True:
                    continue
                
                write_dic_row(p_array[-3], p_array[-2], p_array[-1], end_path_str, file_id, "Replace")
                file_count += 1

            # delete operator
            for file_delete in root.findall('.//leaf[@operation="delete"]'):
                file_id = file_delete.attrib['ID']
                
                remove_dic_row(file_id)
                    
                file_count -= 1
                del_file_count += 1

    for dos_key in dic:
        for i, file_path in enumerate(dic[dos_key]['FilePath']):
            region = dic[dos_key]['Region'][i]
            dos_id = dic[dos_key]['Dossier_ID'][i]
            seq_no = dic[dos_key]['Seq_No'][i]
            file_path = dic[dos_key]['FilePath'][i]
            file_id = dic[dos_key]['FileID'][i]
            operation = dic[dos_key]['Operation'][i]
            write_excel_row(worksheet, row, region, dos_id, seq_no, file_path, file_id, operation)

            row += 1 

            seq_key = return_sequence_key(file_path)
            if seq_key not in sequences:
                sequences.append(seq_key)
                seq_cnt += 1 

    print("\nFinished 1/3 - Extracting the relevant file output paths from the index and regional .xmls...")

    print("\nTotal file count:", file_count)
    print("Deleted file count:", del_file_count)
    print("Sequences count:", seq_cnt)
    print("\nWriting eCTD_XML_Report.xlsx...")
    workbook.close()

# extracting the current actual eCTD output from backups of myPress folders
def extract_file_paths(ectd_path):
    file_workbook = X2.Workbook('FilePaths.xlsx')
    worksheet2 = file_workbook.add_worksheet("File_Paths")
    headers2 = ["Region", "Dossier Identifier", "Sequence number", "File path"]
    header_format_2 = file_workbook.add_format({'bold': True})
    for i, header in enumerate(headers2):
        worksheet2.write(0, 0 + i, header, header_format_2)
    worksheet2.freeze_panes(1, 0)

    row = 1
    file_count = 0

    pathlist = Path(ectd_path).glob("**/*.*")

    for path in pathlist:
        pp = PureWindowsPath(path)

        # exclude all backup folders of any kind
        if 'ackup' in str(pp):
            continue

        test_array = str(pp).split('\\')

        for region in allowed_regions:
            for path_element in test_array:
                if region == path_element:
                    if('util' or 'evalidator') in str(pp):
                        continue
                
                    dos_key = return_dos_key(str(pp), True)

                    if dos_key == 'false':
                        continue
                
                    if dos_key not in file_dic:
                        file_dic[dos_key] = {
                        'Region': [],
                        'Dossier_ID': [],
                        'Seq_No': [],
                        'FilePath': []
                    }
                    file_path_str = assemble_file_path(str(pp))
                    file_dic[dos_key]['FilePath'].append(file_path_str)

    for seq_key in file_dic:
        for i, file_path in enumerate(file_dic[seq_key]['FilePath']):
            file_path_array = file_dic[seq_key]['FilePath'][i].split("/")

            if 'pdf' in file_path_array[-1] or 'doc' in file_path_array[-1] or 'docx' in file_path_array[-1]:
                # for some reason, my path report from MS was ending with \n. I mitigated it by removing it from the string, to avoid any problems in the final comparison program
                new_path = file_path.replace("\n", "")

                file_dic[seq_key]['Region'].append(file_path_array[0])
                file_dic[seq_key]['Dossier_ID'].append(file_path_array[1])
                file_dic[seq_key]['Seq_No'].append(file_path_array[2])

                file_dic[seq_key]['FilePath'][i] = new_path

                worksheet2.write(row, 0, file_path_array[0])
                worksheet2.write(row, 1, file_path_array[1])
                worksheet2.write(row, 2, file_path_array[2])
                worksheet2.write(row, 3, new_path)

                row += 1
                file_count += 1

    #closing workbook
    print("\nFinished 2/3 - Extracting the relevant file output paths from the defined folder...")
    print("\nTotal file count:", file_count)
    print("Writing FilePaths.xlsx...")
    file_workbook.close()

def compare_paths():
    # missing files workbook¸¸
    workbook_3 = X3.Workbook('eCTD_file_compare.xlsx')
    worksheet3 = workbook_3.add_worksheet("eCTD_File_Compare")
    headers3 = ["Region", "Dossier Identifier", "Sequence number", "File path", "File ID"]
    header_format1 = workbook_3.add_format({'bold': True})   
    for i, header in enumerate(headers3):
        worksheet3.write(0, 0 + i, header, header_format1)
    worksheet3.freeze_panes(1, 0)

    row_1 = 1

    # missing sequences Excel report
    workbook_4 = X4.Workbook('eCTD_sequence_compare.xlsx')
    worksheet2 = workbook_4.add_worksheet("eCTD_File_Compare")
    headers2 = ["Region", "Dossier Identifier", "Sequence number"]
    header_format2 = workbook_4.add_format({'bold': True})
    for i, header in enumerate(headers2):
        worksheet2.write(0, 0 + i, header, header_format2)
    worksheet2.freeze_panes(1, 0)

    row_2 = 1

    sequences_affected = []
    files_affected = []

    for id_1 in file_dic:
        for id_2 in dic:
            if id_1 == id_2:
                for j, item_2 in enumerate(dic[id_1]['FilePath']):
                    found = False
                    for k, item_1 in enumerate(file_dic[id_1]['FilePath']):
                        if item_1 == item_2:
                            found = True        

                    if found == False:
                        seq_key = return_sequence_key(dic[id_1]['FilePath'][j])
                        if seq_key not in sequences_affected:
                            sequences_affected.append(seq_key)

                            worksheet2.write(row_2, 0, dic[id_1]['Region'][j])
                            worksheet2.write(row_2, 1, dic[id_1]['Dossier_ID'][j])
                            worksheet2.write(row_2, 2, dic[id_1]['Seq_No'][j])
                            row_2 += 1                           

                        if dic[id_1]['FilePath'][j] not in files_affected:
                            worksheet3.write(row_1, 0, dic[id_1]['Region'][j])
                            worksheet3.write(row_1, 1, dic[id_1]['Dossier_ID'][j])
                            worksheet3.write(row_1, 2, dic[id_1]['Seq_No'][j])
                            worksheet3.write(row_1, 3, dic[id_1]['FilePath'][j])
                            worksheet3.write(row_1, 4, dic[id_1]['FileID'][j])

                            files_affected.append(dic[id_1]['FilePath'][j])
                            row_1 += 1

    print("\nFinished 3/3 - Comparing the real eCTD file output vs relevant file output paths from the index and regional .xmls...")

    affected_files_cnt = row_1 - 1
    print("\nAffected files count:", affected_files_cnt)
    affected_seq_cnt = row_2 - 1
    print("Affected sequences count:", affected_seq_cnt)

    print("\nCreating eCTD_sequence_compare.xlsx...")
    print("Creating eCTD_file_compare.xlsx...\n")

    #closing workbook
    workbook_4.close()
    workbook_3.close()

# the main part of the program

allowed_regions = []

region_str = input("Enter number of regions that you want to check the missing files for:\n")

if region_str.isnumeric == False:
    print("Please enter a valid number number that is larger than 0.")

region_num = int(region_str)
if region_num < 0:
    print("Please enter a valid number number that is larger than 0.")

i = 0

while i < region_num:
    input_region = input("Enter the folder name of the region from eCTD output folder:\n")
    if input_region.isnumeric == False:
        print("Error - Please input the name of the region only.")           

    allowed_regions.append(input_region)
    i += 1

output_folder_path = input("Please enter a path to check and compare existing eCTD output files vs .xml definition of the sequences:\n")

# Extracting the golden standard definition written in index.xml and m1/regional.xml into eCTD_XML_Report.xlsx Excel file
extract_paths_from_xml(output_folder_path)

# Extracting the current folder output of .pdf, .doc. and .docx files from the defined folder path
extract_file_paths(output_folder_path)

# Comparing the results generated by the two previous sequences and pinpointing missing files
compare_paths()

os.system("pause")