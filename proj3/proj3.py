import sys
print(sys.version)
import streamlit as st
import subprocess
from datetime import datetime

# sets the title of the page
st.set_page_config(page_title="Team: Madhur and Siddharth")

# sets the session to prevent re-rendering of the whole component.
if 'replace' not in st.session_state:
    st.session_state['replace'] = False

if "filter" not in st.session_state:
    st.session_state['filter'] = False


st.header("Welcome, this is GUI for Project 3.")
st.subheader("And this is team Madhur n Siddharth.")

def Take_Input():
    try:
        # taking inputs for constant_fk2d, multiplying_factor, Shear_velocity
        st.markdown("---")
        st.markdown("<h1 style='text-align: center;'>Project PSAT v3.0</h1>", unsafe_allow_html=True)
        st.markdown("---")
        constant_fk2d = st.number_input('Enter constant_fk2d : ', step=.01, format="%.2f")

        multiplying_factor = st.number_input('Enter multiplying_factor : ', step=.01, format="%.2f")

        Shear_velocity = st.number_input('Enter Shear_velocity : ', step=.01, format="%.2f")

        # initializing variables to be used
        corr, snr, k_value, lambda_value = 0, 0, 0.0, 0.0

        # button and logic for filtering written
        if st.button("Go to Filter") or st.session_state.filter:
            st.session_state['filter'] = True
            if constant_fk2d>0 and multiplying_factor>0 and Shear_velocity>0:
                st.text(" 1. C\n 2. S\n 3. A\n 4. C & S\n 5. C & A\n 6. S & A\n 7. C & S & A\n 8. All Combine")
                option = st.selectbox("Choose the filtering method from above : ", options=('Select here',1,2,3,4,5,6,7,8), key='option')
                if option == 1:
                    corr = st.number_input('Enter Threshold value of C : ', step=1)
                elif option == 2:
                    snr = st.number_input('Enter Threshold value of S : ', step=1)
                elif option == 3:
                    lambda_value = st.number_input('Enter Lambda value of A : ', step=.01, format="%.2f")
                    k_value = st.number_input('Enter k value of A : ', step=.01, format="%.2f")
                elif option == 4:
                    corr = st.number_input('Enter Threshold value of C : ', step=1)
                    snr = st.number_input('Enter Threshold value of S : ', step=1)
                elif option == 5:
                    corr = st.number_input('Enter Threshold value of C : ', step=1)
                    lambda_value = st.number_input('Enter Lambda value of A : ', step=.01, format="%.2f")
                    k_value = st.number_input('Enter k value of A  : ', step=.01, format="%.2f")
                elif option == 6:
                    snr = st.number_input('Enter Threshold value of S :', step=1)
                    lambda_value = st.number_input('Enter Lambda value of A : ', step=.01, format="%.2f")
                    k_value = st.number_input('Enter k value of A : ', step=.01, format="%.2f")
                elif option == 7 or option == 8:
                    corr = st.number_input('Enter Threshold value of C : ', step=1)
                    snr = st.number_input('Enter Threshold value of S : ', step=1)
                    lambda_value = st.number_input('Enter Lambda value of A : ', step=.01, format="%.2f")
                    k_value = st.number_input('Enter k value of A : ', step=.01, format="%.2f")
                if st.button("Go to Replace") or st.session_state.replace:
                    if option>0:
                        st.session_state['replace'] = True
                        replacement_method = st.selectbox("Choose Replacement Method From Below : ", ['1. Previous Point', '2. 2*last-2nd_last', '3. Overall Mean', '4. 12_Point_Strategy', '5. Mean of Previous 2 points', '6. All Sequential', '7. All Parallel'])
                        replacement_method=int(replacement_method[0])
                        if st.button("Compute"):
                                if replacement_method>0:
                                    start_time = datetime.now()
                                    with st.spinner('Computing...'):
                                        subprocess.run(["python", "psat_v3.py", str(constant_fk2d), str(multiplying_factor), str(Shear_velocity), str(option), str(corr), str(snr), str(lambda_value), str(k_value), str(replacement_method)])
                                    end_time = datetime.now()
                                    st.markdown("<h3>Computed! You can check your Results_v2.csv to see the computed values!!</h3>", unsafe_allow_html=True)
                                    st.markdown("<h2>Run time</h2>", unsafe_allow_html=True)
                                    st.write(f'Start time : {start_time.strftime("%c")}')
                                    st.write(f'End time : {end_time.strftime("%c")}')
                                    st.write(f'Duration : {end_time - start_time}')
    except:
        st.text("Something went wrong. Please refresh and try again!!")
        exit(1)
Take_Input()