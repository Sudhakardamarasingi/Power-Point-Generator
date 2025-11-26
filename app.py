import streamlit as st
import requests
import subprocess
st.title("AGENTIC BASED POWER POINT GENERATOR")

prompt = st.text_area("please write the details of how to create the ppt")

if st.button("Get PPT"):
  if prompt:
    requests.post(url="https://sudha-mad-max-1997.app.n8n.cloud/webhook/f4892281-e1a0-429c-ae0a-16661a18e576",json={"prompt":prompt})
    
    if response.status_code==200:
      st.write("success")
      
      with open("app1.py","w") as file:
        file.write(response.json()["output"])
      subprocess.run(["python","app1.py"])

with  open("app1.py","rb") as f1:
st.download_button(
    label = "Download PPT",
    data = f1,
    file_name = "data.pptx")
