import ClointFusion as cf

def copyToExcel(headerName,excel_details,li2):
    # ---------------------------create excel file-------------------
    NewFilePath = cf.gui_get_any_input_from_user(msgForUser='Enter The File Path(without quotes) Where You want to save its copy', password=False, multi_line=False, mandatory_field=True)
    NewFileName = cf.gui_get_any_input_from_user(msgForUser="Enter the excel filename without extension")
    cf.excel_create_excel_file_in_given_folder(fullPathToTheFolder=NewFilePath,excelFileName= NewFileName)
    
    # -------------------------getting path of new file--------------------
    b="\\"+NewFileName+".xlsx"
    finalPath = NewFilePath+b
    
    # ----------------------copying to a new excel file---------------------
    for q in range(0,len(headerName)):
        for p in range(0,len(li2)):
            cf.excel_set_single_cell(excel_path=finalPath,sheet_name=excel_details[1],columnName=headerName[q],cellNumber=p,setText=li2[p][q])

options = None
excel_details = None

options = cf.gui_get_dropdownlist_values_from_user(msgForUser='Select any option', dropdown_list=['Copy Full Excel file', 'Other Operations'], multi_select=False)

if (options[0] == 'Copy Full Excel file'):

    # ------------------------getting which excel to copy----------------------
    excel_details = cf.gui_get_excel_sheet_header_from_user(msgForUser='Select Excel File')
    headerName = cf.excel_get_all_header_columns( excel_path=excel_details[0], sheet_name=excel_details[1])

    # ------------------------Row,column count--------------------------
    count = cf.excel_get_row_column_count(excel_path=excel_details[0], sheet_name=excel_details[1], header=excel_details[2])
    countdup=[count[0],count[1]]

    # ----------------------copy file contents to a list----------------------
    li2=[]
    if(excel_details[2]==0):    
        li2 = cf.excel_copy_range_from_sheet(excel_path=excel_details[0], sheet_name=excel_details[1], startRow=excel_details[2]+2, startCol=1, endRow=count[0]+excel_details[2], endCol=count[1])
        countdup[0]=count[0]-1
    else:
        li2 = cf.excel_copy_range_from_sheet(excel_path=excel_details[0], sheet_name=excel_details[1], startRow=excel_details[2]+1, startCol=1, endRow=count[0]+excel_details[2], endCol=count[1])
    

    copyToExcel(headerName,excel_details,li2)


else:
    options = None
    excel_details = None
    # ------------------------getting which excel to copy----------------------
    excel_details = cf.gui_get_excel_sheet_header_from_user(msgForUser='Select Excel File')
    headerName = cf.excel_get_all_header_columns( excel_path=excel_details[0], sheet_name=excel_details[1])

    # ------------------------Row,column count--------------------------
    count = cf.excel_get_row_column_count(excel_path=excel_details[0], sheet_name=excel_details[1], header=excel_details[2])
    countdup=[count[0],count[1]]

    # ----------------------copy file contents to a list----------------------
    li2=[]
    if(excel_details[2]==0):    
        li2 = cf.excel_copy_range_from_sheet(excel_path=excel_details[0], sheet_name=excel_details[1], startRow=excel_details[2]+2, startCol=1, endRow=count[0]+excel_details[2], endCol=count[1])
        countdup[0]=count[0]-1
    else:
        li2 = cf.excel_copy_range_from_sheet(excel_path=excel_details[0], sheet_name=excel_details[1], startRow=excel_details[2]+1, startCol=1, endRow=count[0]+excel_details[2], endCol=count[1])

    # --------------------------getting the conditions------------------
    options1 = cf.gui_get_dropdownlist_values_from_user(msgForUser='Select any option', dropdown_list=['greater than the number will be copied','less than the number will be copied','greater than the 1stcharacter will be copied','less than the 1stcharacter will be copied','Copy a particular element'], multi_select=False)
    finalli = []
    t=0

    # ------------greater than the number entered will be added-----------------
    if(options1[0]=='greater than the number will be copied'):
        options = cf.gui_get_dropdownlist_values_from_user(msgForUser='Select any option', dropdown_list=headerName, multi_select=False)
        op=headerName.index(options[0])
        numG = cf.gui_get_any_input_from_user(msgForUser='Enter the number so that all copied elements will be greater than this number ', password=False, multi_line=False, mandatory_field=True)
        try:
            numG = int(numG)
            for j in range(0,countdup[0]):
                if(li2[j][op] > numG):
                    t=1
                    finalli.append(li2[j])
        except:
            print("Please enter only a integer type")
    
    # ------------less than the number entered will be added-----------------
    elif(options1[0]=='less than the number will be copied'):
        options = cf.gui_get_dropdownlist_values_from_user(msgForUser='Select any option', dropdown_list=headerName, multi_select=False)
        op=headerName.index(options[0])
        numG = cf.gui_get_any_input_from_user(msgForUser='Enter the number so that all copied elements will be less than this number ', password=False, multi_line=False, mandatory_field=True)
        try:
            numG = int(numG)
            for j in range(0,countdup[0]):
                if(li2[j][op] < numG):
                    t=1
                    finalli.append(li2[j])
        except:
            print("Please enter only a integer type")

    # ------------greater than the character entered will be added-----------------
    elif(options1[0]=='greater than the 1stcharacter will be copied'):
        options = cf.gui_get_dropdownlist_values_from_user(msgForUser='Select any option', dropdown_list=headerName, multi_select=False)
        op=headerName.index(options[0])
        numG = str(cf.gui_get_any_input_from_user(msgForUser='Enter the first character so that all copied elements will be greater than this character(case sensitive i.e c>C) ', password=False, multi_line=False, mandatory_field=True))
        if(len(numG) == 1):
            for j in range(0,countdup[0]):
                a = ord(numG)
                b = ord(li2[j][op][0])
                if(b > a):
                    t=1
                    finalli.append(li2[j])
        else:
            print("Please enter only one character")

    # ------------less than the character entered will be added-----------------
    elif(options1[0]=='less than the 1stcharacter will be copied'):
        options = cf.gui_get_dropdownlist_values_from_user(msgForUser='Select any option', dropdown_list=headerName, multi_select=False)
        op=headerName.index(options[0])
        numG = str(cf.gui_get_any_input_from_user(msgForUser='Enter the first character so that all copied elements will be less than this character(case sensitive i.e c>C) ', password=False, multi_line=False, mandatory_field=True))
        if(len(numG) ==1):
            for j in range(0,countdup[0]):
                a = ord(numG)
                b = ord(li2[j][op][0])
                if(b < a):
                    t=1
                    finalli.append(li2[j])
        else:
            print("Please enter only one character")

    # ------------particular element entered will be copied-----------------
    elif(options1[0]=='Copy a particular element'):
        numG = str(cf.gui_get_any_input_from_user(msgForUser='Enter the word which you want to copy to your new excel file', password=False, multi_line=False, mandatory_field=True))
        for i in range(0,countdup[0]):
            if numG in li2[i]:
                t=1
                finalli.append(li2[i])
            else:
                try:
                    numG = int(numG)
                    if numG in li2[i]:
                        t=1
                        finalli.append(li2[i])
                except:
                    continue

    if(t == 1):
        copyToExcel(headerName,excel_details,finalli)
    else:
        print("By the above conditions,there is nothing to copy")
