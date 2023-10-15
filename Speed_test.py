import speedtest
import time
import pandas as pd
from datetime import datetime

def check_network_speed():
    while True:
        try:
            st = speedtest.Speedtest()
            download_speed = st.download() / 10**6  
            upload_speed = st.upload() / 10**6  
        except Exception as e:
            print(f"Network Error: {e}")
            download_speed = upload_speed = None

        current_time = datetime.now()
        if download_speed is not None and upload_speed is not None:
            print(f'{time.strftime("%I:%M %p")} - Download Speed: {download_speed:.2f} Mbps, Upload Speed: {upload_speed:.2f} Mbps')

            save_data_to_excel(current_time, download_speed, upload_speed)
        time.sleep(3600)

def save_data_to_excel(timestamp, download_speed, upload_speed):
    try:
        df = pd.read_excel(f"network_speed_{timestamp.date()}.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Time", "Download", "Upload"])

    new_data = {
        "Time": timestamp.strftime("%I:%M %p"),
        "Download": download_speed,
        "Upload": upload_speed
    }

    df = pd.concat([df, pd.DataFrame(new_data, index=[0])], ignore_index=True)
    df.to_excel(f"network_speed_{timestamp.date()}.xlsx", index=False)

if __name__ == "__main__":
    check_network_speed()
