import streamlit as st
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from gspread import Cell
from datetime import date
import datetime
from PIL import Image
import pickle

image = Image.open('jesap.png')

st.set_page_config(page_title='Timetree', page_icon = image, initial_sidebar_state = 'auto')
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive',
         'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']

creds = ServiceAccountCredentials.from_json_keyfile_name('key_conteggio.json', scope)
client = gspread.authorize(creds)


st.markdown('## Business Development Area')
st.markdown('### [Google Sheet Link](https://docs.google.com/spreadsheets/d/1ioR9yXMCJJxuxaiuEBoPKP4-EJLa6RgdYZicrTQpxl0/edit#gid=0)')
link_BD = "https://docs.google.com/spreadsheets/d/1ioR9yXMCJJxuxaiuEBoPKP4-EJLa6RgdYZicrTQpxl0/edit#gid=0"
sht = client.open_by_url(link_BD)

# funzione al click del bottone
def fun():
	st.session_state.multi = []
	if nome != '':
		if not temp_att:
			st.warning('Choose at least one activity')
			return
		double = 0
		double_inprova = 0
		for w in sht.worksheets():
			lower_title = w.title.lower()
			index_inprova = lower_title.find("- trainee")
			if (index_inprova >= 0):
				lower_title = lower_title[0:index_inprova-1]
			lower_name = nome.lower()
			if lower_name in lower_title:
				if (index_inprova >= 0):
					double_inprova += 1
					work_inprova = w
				else:
					double += 1
					work = w
		if double == 0 and double_inprova == 0:
			st.error('No member has been found with this name/surname')
			return
		if double > 1 and (not in_prova):
			st.warning('Multiple members have been found with this name/surname; please try to be more specific.')
			return
		if double_inprova > 1 and (in_prova):
			st.warning('Multiple trainees have been found with this name/surname; please try to be more specific.')
			return
		else:
			if in_prova: 
				if double_inprova == 0:
					st.error('No member matching the search criteria has been found')
					return
				elif double_inprova == 1:
					current_work = work_inprova
			elif (not in_prova):
				if double == 0:
					st.error('No member matching the search criteria has been found')
					return
				elif double == 1:
					current_work = work
		# --- adding elements to google sheet ---
			def next_available_row(worksheet): #funzione
				str_list = list(filter(None, worksheet.col_values(1))) #fa la lista delle colonne del worksheet scegliendo un elemento non vuoto
				return str(len(str_list)+1) #va a prendere il prossimo elemento, che sarà vuoto

			for a in temp_att:
				row = next_available_row(current_work)
				if int(row) > (len(current_work.get_all_values())-1):
					current_work.append_row(["", "", ""])
				c1 = Cell(int(row) , 1, str(data))
				c2 = Cell(int(row) , 2, a)
				c3 = Cell(int(row) , 3, str(dictionary[a]).replace(':', '.'))
				current_work.update_cells([c1,c2,c3], value_input_option='USER_ENTERED')

			st.success(f"{current_work.title}'s Timetree has been updated")
	else:
		st.warning('First name and/or Last name are required')
	return

# --- Interfaccia ----
nome = st.text_input('First name and/or Last name. Check the box if you are a trainee')
in_prova = st.checkbox('Trainee')
data = st.date_input('Date', value=date.today())
options = ['Area Call', 'Monthly Meeting', 'Team', 'Recruiting', 'Mentoring', 'External Project'
,'Internal Project', 'Training', 'HR Buddy Call', 'Case Study', 'Area Organization', 'Task', 'Corporate Event', 'Task Review', 'Business Proposal', 'Quote', 'First Lead Call', 'Follow-up Lead Call', 'Lead Generation', "Lead management", 'Board-Head/Head-Deputy Call', 'Partnership', 'Other']

att = st.multiselect('Activity', options, key="multi")
dictionary = {}
temp_att = []
for a in att:
	if a == "External Project":
		### estrazione nomi progetti in corso
		prog_link = "https://docs.google.com/spreadsheets/d/1kaiBTPxp-o0IVn1j54QGuPJOeYdHxJ9Iqg30QwYqnGI/edit#gid=1965451645"
		prog_spread_sht = client.open_by_url(prog_link)
		prog_sht = prog_spread_sht.get_worksheet(1)

		column_b = prog_sht.col_values(1)  # Column B (poi diventata colonna A con nuovo foglio) is index 1
		column_d = prog_sht.col_values(3)  # Column D (poi diventata colonna C con nuovo foglio) is index 3
		column_e = prog_sht.col_values(4)

		progetti_in_corso = [] #era vuoto []
		for name, value, state in zip(column_b, column_d, column_e):
			if value.lower() == 'in corso' and state.lower() == 'progetto esterno':
				progetti_in_corso.append(name)
		### fine estrazione
		sel_prog = st.selectbox("Select one project", progetti_in_corso)
		if sel_prog:
			temp_att.append('External Project - ' + sel_prog)
			n_ore = st.time_input(f'Number of hours worked: {sel_prog}', datetime.time(1, 0), key=sel_prog+'1')
			dictionary['External Project - ' + sel_prog] = n_ore

	else:
		temp_att.append(a)
		n_ore = st.time_input(f'Number of hours worked: {a}', datetime.time(1, 0), key=a)
		dictionary[a] = n_ore

data = data.strftime("%d/%m/%Y")
sub = st.button("Submit",on_click=fun)

# Font: Nunito, colore bottone blu
m = st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@200;300;400&display=swap');
html, body, [class*="css"]  {
	font-family: 'Nunito', sans-serif; 
	text-color: "#ffffff";
}
div.stButton > button:first-child {
    background-color: #2e9aff;
    border-color: #2e9aff;
}

div.stButton > button:hover {
    color:#ffffff;
	background-color: rgba(89, 55, 146, 0.8);
	border-color: #ffffff;
    }
				
header.css-1avcm0n, section {
	background-color: rgba(89, 55, 146, 0.8);
}
</style>""", unsafe_allow_html=True)
