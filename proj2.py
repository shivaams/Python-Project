import streamlit as st
import pandas as pd
import os
import tut07 as tut
from Navigiation_Bar import nav_bar as nb
from DownloadButton import download_button as db  
from zipfile import ZipFile  

# -------------------------------------------------------------------------------------------------------------------
def downloading(file,mod,download_buttons,tables,names):

    count = 0
    if mod == "":
        excel = tut.octant_analysis(name=file)
        count = 1
    else:
        if int(mod)>0:
            excel = tut.octant_analysis(name=file,mod=int(mod))
            count = 1
    
    if count:
        
        with open(excel, "rb") as ex:
            s = ex.read()
            table = pd.read_excel(excel)
            tables.append(table)
            download_button_str = db(s, excel, f'Download "{excel[7:-5]}"')
            download_buttons.append(download_button_str)
            names.append(excel)

# -------------------------------------------------------------------------------------------------------------------
def zip_creator(names):

    zip_file = "Output.zip"

    with ZipFile(zip_file,'w') as zipf:
        for name in names:
            zipf.write(name)

    return zip_file
        
# -------------------------------------------------------------------------------------------------------------------     
def checker(mod):

    if mod == "":
        return 1
    else:
        if int(mod)>0:
            return 1
        else:
            try:
                st.error('Enter a valid mod value!', icon="🚨")
                return 0
            except:
                pass
            
    return 1

# -------------------------------------------------------------------------------------------------------------------
def page_style():

    st.markdown("---")
    col1,col2 = st.columns([2,5])
    with col1:
        st.write("#")
        st.write("Enter Mod value")
    with col2:
        mod = st.number_input("a",label_visibility="hidden",min_value=50,value=50,step=1)
    st.markdown("---")

    return mod

# -------------------------------------------------------------------------------------------------------------------
def selected_files_converter(files):

    mod = page_style()

    convert = st.button("Convert")
    download_buttons = []
    tables = []
    names = []

    

    if convert and checker(mod):
        download_buttons.clear()
        if len(files):
            status = 1
            pb = st.progress(0)
            mul = 0
            
            for file in files:
                mul = (status/len(files))
                pb.progress(mul)
                status += 1
                if file is not None:
                    downloading(file,mod,download_buttons,tables,names)
            pb.progress(mul)
            pb.empty()
            
        else:
            st.error('Please select an excel file!', icon="🚨")

    if len(download_buttons):
        st.subheader("Results:")

        # Code to donwload Zip file
        zip_file = zip_creator(names)

        with open (zip_file,"rb") as zf:
            z_file = zf.read()
            download_button_str = db(z_file,zip_file,f'Download zip file')
            st.markdown(download_button_str,unsafe_allow_html=True)
        
        tab_list = []
        file_n = "File "
        c = 1
        for (button,table,name) in zip(download_buttons,tables,names):
            
            tab_list.append(file_n+str(c))
            c+=1
        
        tab_list = st.tabs(tab_list)

        for i in range(len(tab_list)):
            
            col1,col2 = tab_list[i].columns(2)
            table = tables[i].style.highlight_null(props="color: transparent;")
            with col1:
                tab_list[i].write("File: "+names[i])
            with col2:
                tab_list[i].markdown(download_buttons[i], unsafe_allow_html=True)
            
            tab_list[i].dataframe(table)
            
# -------------------------------------------------------------------------------------------------------------------                          
def batch_files_converter(path):
    mod = page_style()

    convert = st.button("Convert")

    download_buttons = []
    tables = []
    names = []
    if convert:
        download_buttons.clear()
        files = []

        if os.path.exists(path) == 0:
            st.error('Please enter a valid path!', icon="🚨")

        else:
            for filename in os.listdir(path):
                ext = os.path.splitext(filename)[1]
                if ext == ".xlsx":
                    file = os.path.join(path,filename)
                    files.append(file)

            if len(files) == 0:
                st.error('Please enter a valid path!', icon="🚨")
            else:
                status = 1
                pb = st.progress(0)
                mul = 0
                for file in files:
                    mul = (status/len(files))
                    pb.progress(mul)
                    status += 1
                    downloading(file,mod,download_buttons,tables,names)
                pb.progress(mul)
                pb.empty()


    # if len(download_buttons):
    #     st.subheader("Results:")

    #     zip_file = zip_creator(names)

    #     with open (zip_file,"rb") as zf:
    #         z_file = zf.read()
    #         download_button_str = db(z_file,zip_file,f'Download zip file')
    #         st.markdown(download_button_str,unsafe_allow_html=True)
        

    #     st.markdown("---")

    #     for (button,table,name) in zip(download_buttons,tables,names):
            
    #         name = name[7:-5]
            
    #         col1,col2 = st.columns(2)
    #         table = table.style.highlight_null(props="color: transparent;")
            
    #         with col1:
    #             st.write("File: "+name)
    #         with col2:
    #             st.markdown(button, unsafe_allow_html=True)

    #         st.dataframe(table)
    #         st.markdown("---")
    if len(download_buttons):
        st.subheader("Results:")

        # Code to donwload Zip file
        zip_file = zip_creator(names)

        with open (zip_file,"rb") as zf:
            z_file = zf.read()
            download_button_str = db(z_file,zip_file,f'Download zip file')
            st.markdown(download_button_str,unsafe_allow_html=True)
        
        tab_list = []
        file_n = "File "
        c = 1
        for (button,table,name) in zip(download_buttons,tables,names):
            
            tab_list.append(file_n+str(c))
            c+=1
        
        tab_list = st.tabs(tab_list)

        for i in range(len(tab_list)):
            
            col1,col2 = tab_list[i].columns(2)
            table = tables[i].style.highlight_null(props="color: transparent;")
            with col1:
                tab_list[i].write("File: "+names[i])
            with col2:
                tab_list[i].markdown(download_buttons[i], unsafe_allow_html=True)
            
            tab_list[i].dataframe(table)

# -------------------------------------------------------------------------------------------------------------------
#Help
def proj_octant_gui():
	st.set_page_config(page_title="Project")

	with open('style.css') as f:
		st.markdown(f'<style>{f.read()}</style>',unsafe_allow_html=True)

	selected = nb()


	st.title("Project 2")

	if selected == "File Selection":
		files = st.file_uploader("Upload files",accept_multiple_files=True)
		selected_files_converter(files)

	if selected == "Path":
		path = st.text_input("Enter path to the folder for bulk conversion")
		batch_files_converter(path)

# -------------------------------------------------------------------------------------------------------------------
###Code


proj_octant_gui()
