import pandas as pd
import streamlit as st
from streamlit_option_menu import option_menu
import sqlite3
from passlib.hash import md5_crypt

# DB Management
conn = sqlite3.connect('data.db')
c = conn.cursor()

#c.execute('DROP FROM userstable WHERE username="ass"')


def create_usertable():
    c.execute('CREATE TABLE IF NOT EXISTS userstable(username TEXT, password TEXT)')


def add_userdata(username, password):
    c.execute('INSERT INTO userstable(username, password) VALUES (?,?)',
              (username, md5_crypt.hash(password)))
    conn.commit()


def delete_userdata(username):
    c.execute('DELETE FROM userstable WHERE username=?',
              (username,))
    # f.encrypt(str.encode(password))
    conn.commit()


def login_user(username):
    c.execute('SELECT * FROM userstable WHERE username=?',
              (username,))
    # f.encrypt(str.encode(password))
    data = c.fetchone()
    return data


def view_all_users():
    c.execute('SELECT * FROM userstable')
    data = c.fetchall()
    return data


def admin():
    c.execute('SELECT * FROM userstable WHERE username="admin"')
    data = c.fetchone()
    return data


def main():

    st.set_page_config(page_title='Factory Consolidate View',
                       page_icon='Dell_Logo_Blue_rgb.png', layout="wide")

    st.header('GPP Workplan Consolidate View')

    menu = ["Login", "Sign Up"]

    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "Login":
        username = st.sidebar.text_input("User Name")
        password = st.sidebar.text_input("Password", type='password')
        if st.sidebar.checkbox("Login"):
            # if password == "123":
            create_usertable()
            # md5_crypt.verify(password, login_user(username))
            result = md5_crypt.verify(password, login_user(username)[1])
            # result = login_user(username)  # , password)

            if result:
                st.success("Logged in as {}".format(username))

                st.markdown("""

                    <style>
                    .css-9s5bis {
                            color: transparent;
                        }

                    </style>

                    """, unsafe_allow_html=True)

                st.markdown('<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">', unsafe_allow_html=True)
                excel_file = r'Consolidated Factory Workplan.xlsx'
                sheet_name = 'Workplans'

                # user_result = view_all_users()

                # clean_db = pd.DataFrame(user_result, columns=[
                #                         "Username", "Password"])
                # st.dataframe(clean_db) #view database

                df = pd.read_excel(excel_file,
                                   sheet_name=sheet_name,
                                   usecols='B:K',
                                   skiprows=(0, 1, 2, 3, 4, 5),
                                   header=None)

                df1 = df.fillna('')
                time = pd.read_excel(excel_file,
                                     sheet_name=sheet_name,
                                     usecols='F:F',
                                     nrows=1,
                                     header=3)

                time1 = time.iloc[0][0]

                st.subheader('Update On : ' + time1)

                st.markdown("""
                <nav class="navbar fixed-top navbar-expand-lg navbar-dark" style="background-color: #3498DB; font:serif;">
                <a class="navbar-brand" style="max-width: 500px;
                margin: auto;">
                <img  src ="media/ece2b310c7c817dd3a7e5da29a1adb5608ea6b51036e759b57105c52.png" alt="" width="30" height="30">
                    Dell Global Factory Workplan
                </a>
                <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNav" aria-controls="navbarNav" aria-expanded="false" aria-label="Toggle navigation">
                    <span class="navbar-toggler-icon"></span>
                </button>
                <div class="collapse navbar-collapse" id="navbarNav">
                </div>
                </nav>
                """, unsafe_allow_html=True)

                with open('style.css') as f:
                    st.markdown(f'<style>{f.read()}</style>',
                                unsafe_allow_html=True)

                a, b, c, d, e, f, g, h = st.columns((1, 1, 1, 1, 1, 1, 1, 1))

                with a:
                    group = option_menu(
                        menu_title=None,
                        options=['ALL'],
                        default_index=0,
                        menu_icon="building",
                        icons=["building"],
                        styles={
                            "menu_icon": {"color": "#3498DB"},
                            "nav-link-selected": {"background-color": "#3498DB"}
                        },
                        orientation="horizontal"
                    )
                    group1 = st.checkbox('All', key='group1')

                with b:
                    group = option_menu(
                        menu_title=None,
                        options=['BRH1'],
                        menu_icon="building",
                        icons=['bank2'],
                        styles={
                            "menu_icon": {"color": "#3498DB"},
                            "nav-link-selected": {"background-color": "#3498DB"}
                        },
                        orientation="horizontal"

                    )
                    group2 = st.checkbox(
                        'BRH1', key='group_2', disabled=(group1 is True))
                with c:
                    group = option_menu(
                        menu_title=None,
                        options=['EMFP'],
                        menu_icon="building",
                        icons=['bank2'],
                        styles={
                            "menu_icon": {"color": "#3498DB"},
                            "nav-link-selected": {"background-color": "#3498DB"}
                        },
                        orientation="horizontal"

                    )
                    group8 = st.checkbox('EMFP', disabled=(group1 is True))
                with d:
                    group = option_menu(
                        menu_title=None,
                        options=['APCC'],
                        menu_icon="building",
                        icons=['bank2'],
                        styles={
                            "menu_icon": {"color": "#3498DB"},
                            "nav-link-selected": {"background-color": "#3498DB"}
                        },
                        orientation="horizontal"

                    )
                    group7 = st.checkbox('APCC', disabled=(group1 is True))
                with e:
                    group = option_menu(
                        menu_title=None,
                        options=['ICC'],
                        menu_icon="building",
                        icons=['bank2'],
                        styles={
                            "menu_icon": {"color": "#3498DB"},
                            "nav-link-selected": {"background-color": "#3498DB"}
                        },
                        orientation="horizontal"

                    )
                    group6 = st.checkbox('ICC', disabled=(group1 is True))
                with f:
                    group = option_menu(
                        menu_title=None,
                        options=['CCC2'],
                        menu_icon="building",
                        icons=['bank2'],
                        styles={
                            "menu_icon": {"color": "#3498DB"},
                            "nav-link-selected": {"background-color": "#3498DB"}
                        },
                        orientation="horizontal"

                    )
                    group3 = st.checkbox('CCC2', disabled=(group1 is True))
                with g:
                    group = option_menu(
                        menu_title=None,
                        options=['CCC4'],
                        menu_icon="building",
                        icons=['bank2'],
                        styles={
                            "menu_icon": {"color": "#3498DB"},
                            "nav-link-selected": {"background-color": "#3498DB"}
                        },
                        orientation="horizontal"

                    )
                    group4 = st.checkbox('CCC4', disabled=(group1 is True))
                with h:
                    group = option_menu(
                        menu_title=None,
                        options=['CCC6'],
                        menu_icon="building",
                        icons=['bank2'],
                        styles={
                            "menu_icon": {"color": "#3498DB"},
                            "nav-link-selected": {"background-color": "#3498DB"}
                        },
                        orientation="horizontal"

                    )
                    group5 = st.checkbox('CCC6', disabled=(group1 is True))

                hide_st_style = """ 
                                <style>
                                #MainMenu {visibility: hidden;}
                                header {visibility: hidden;}
                                footer {visibility: hidden;}
                                </style>
                                """

                st.markdown(hide_st_style, unsafe_allow_html=True)

                # -- LOAD DATAFRAME

                ccc4 = df1.iloc[0:9, :].astype(str)
                ccc2 = df1.iloc[9:18, :].astype(str)
                ccc6 = df1.iloc[17:21, :].astype(str)
                apcc = df1.iloc[25:33, :].astype(str)
                emfp = df1.iloc[41:44, :].astype(str)
                brh1 = df1.iloc[49:58, :].astype(str)
                icc = df1.iloc[33:41, :].astype(str)

                cv, cn, cm = st.columns((1, 4, 1))
                c1, c2, c3, c4 = st.columns((2, 2, 2, 2))
                c5, c6, c7 = st.columns((2, 2, 2))
                # cc,cb,ck = st.columns((1,4,1))

                if group1:

                    if group1:

                        with c1:
                            with open('style3.css') as f:
                                st.markdown(f'<style>{f.read()}</style>',
                                            unsafe_allow_html=True)

                            D2 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=14)
                            t2 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=14)

                            DD2 = D2.iloc[0][0]
                            tt2 = t2.iloc[0][0]
                            st.write('Factory: ' + tt2)
                            st.write('Date: ' + DD2)
                            st.write(ccc2)

                        with c2:
                            with open('style3.css') as f:
                                st.markdown(f'<style>{f.read()}</style>',
                                            unsafe_allow_html=True)
                            D3 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=5)
                            t3 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=5)

                            DD3 = D3.iloc[0][0]
                            tt3 = t3.iloc[0][0]
                            st.write('Factory: ' + tt3)
                            st.write('Date: ' + DD3)
                            st.write(ccc4)

                        with c3:
                            with open('style3.css') as f:
                                st.markdown(f'<style>{f.read()}</style>',
                                            unsafe_allow_html=True)
                            D4 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=22)
                            t4 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=22)

                            DD4 = D4.iloc[0][0]
                            tt4 = t4.iloc[0][0]
                            st.write('Factory: ' + tt4)
                            st.write('Date: ' + DD4)
                            st.write(ccc6)

                        with c5:
                            with open('style3.css') as f:
                                st.markdown(f'<style>{f.read()}</style>',
                                            unsafe_allow_html=True)
                            D5 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=30)
                            t5 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=30)

                            DD5 = D5.iloc[0][0]
                            tt5 = t5.iloc[0][0]
                            st.write('Factory: ' + tt5)
                            st.write('Date: ' + DD5)
                            st.write(apcc)

                        with c6:
                            with open('style3.css') as f:
                                st.markdown(f'<style>{f.read()}</style>',
                                            unsafe_allow_html=True)
                            D7 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=46)
                            t7 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=46)

                            DD7 = D7.iloc[0][0]
                            tt7 = t7.iloc[0][0]
                            st.write('Factory: ' + tt7)
                            st.write('Date: ' + DD7)
                            st.write(emfp)

                        with c7:
                            with open('style3.css') as f:
                                st.markdown(f'<style>{f.read()}</style>',
                                            unsafe_allow_html=True)
                            D6 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=38)
                            t6 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=38)

                            DD6 = D6.iloc[0][0]
                            tt6 = t6.iloc[0][0]
                            st.write('Factory: ' + tt6)
                            st.write('Date: ' + DD6)
                            st.write(icc)

                        with c4:
                            with open('style3.css') as f:
                                st.markdown(f'<style>{f.read()}</style>',
                                            unsafe_allow_html=True)
                            D1 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=54)
                            t1 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=54)

                            DD1 = D1.iloc[0][0]
                            tt1 = t1.iloc[0][0]
                            st.write('Factory: ' + tt1)
                            st.write('Date: ' + DD1)
                            st.write(brh1)

                else:

                    if group2:
                        with open('style1.css') as f:
                            st.markdown(f'<style>{f.read()}</style>',
                                        unsafe_allow_html=True)
                        with cn:
                            D1 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=54)
                            t1 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=54)

                            DD1 = D1.iloc[0][0]
                            tt1 = t1.iloc[0][0]
                            st.write('Factory: ' + tt1)
                            st.write('Date: ' + DD1)
                            st.write(brh1)

                    if group3:
                        with open('style1.css') as f:
                            st.markdown(f'<style>{f.read()}</style>',
                                        unsafe_allow_html=True)
                        with cn:
                            D2 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=14)
                            t2 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=14)

                            DD2 = D2.iloc[0][0]
                            tt2 = t2.iloc[0][0]
                            st.write('Factory: ' + tt2)
                            st.write('Date: ' + DD2)
                            st.write(ccc2)

                    if group4:
                        with open('style1.css') as f:
                            st.markdown(f'<style>{f.read()}</style>',
                                        unsafe_allow_html=True)
                        with cn:
                            D3 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=5)
                            t3 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=5)

                            DD3 = D3.iloc[0][0]
                            tt3 = t3.iloc[0][0]
                            st.write('Factory: ' + tt3)
                            st.write('Date: ' + DD3)
                            st.write(ccc4)

                    if group5:
                        with open('style1.css') as f:
                            st.markdown(f'<style>{f.read()}</style>',
                                        unsafe_allow_html=True)
                        with cn:
                            D4 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=22)
                            t4 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=22)

                            DD4 = D4.iloc[0][0]
                            tt4 = t4.iloc[0][0]
                            st.write('Factory: ' + tt4)
                            st.write('Date: ' + DD4)
                            st.write(ccc6)

                    if group7:
                        with open('style1.css') as f:
                            st.markdown(f'<style>{f.read()}</style>',
                                        unsafe_allow_html=True)
                        with cn:
                            D5 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=30)
                            t5 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=30)

                            DD5 = D5.iloc[0][0]
                            tt5 = t5.iloc[0][0]
                            st.write('Factory: ' + tt5)
                            st.write('Date: ' + DD5)
                            st.write(apcc)

                    if group6:
                        with open('style1.css') as f:
                            st.markdown(f'<style>{f.read()}</style>',
                                        unsafe_allow_html=True)
                        with cn:
                            D6 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=38)
                            t6 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=38)

                            DD6 = D6.iloc[0][0]
                            tt6 = t6.iloc[0][0]
                            st.write('Factory: ' + tt6)
                            st.write('Date: ' + DD6)
                            st.write(icc)

                    if group8:
                        with open('style1.css') as f:
                            st.markdown(f'<style>{f.read()}</style>',
                                        unsafe_allow_html=True)
                        with cn:
                            D7 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='J:J',
                                               nrows=1,
                                               header=46)
                            t7 = pd.read_excel(excel_file,
                                               sheet_name=sheet_name,
                                               usecols='D:D',
                                               nrows=1,
                                               header=46)

                            DD7 = D7.iloc[0][0]
                            tt7 = t7.iloc[0][0]
                            st.write('Factory: ' + tt7)
                            st.write('Date: ' + DD7)
                            st.write(emfp)

                st.write(
                    'To download the full page of consolidate view, click download button below :')
                btn = st.download_button(
                    label='Download File',
                    data=excel_file,
                    file_name=excel_file,
                )
            else:
                st.warning("Incorrect Username or password")

    elif choice == "Sign Up":
        st.subheader("Create New Account")
        admin_code = st.text_input("Admin Code")  # asdkjh12390u!(9KLmSDJ
        new_user = st.text_input("Username")
        new_password = st.text_input("Password", type='password')

        if st.button("Sign Up"):

            if md5_crypt.verify(admin_code, admin()[1]):
                create_usertable()
                add_userdata(new_user, new_password)
                st.success("You have successfully created a valid account")
                st.info("Go to Login Menu to login")
            else:
                st.warning("Wrong admin code.")

    st.markdown("""
    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
    """, unsafe_allow_html=True)


if __name__ == '__main__':
    main()
