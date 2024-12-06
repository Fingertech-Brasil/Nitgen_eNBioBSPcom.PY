import comtypes.client
import logging
import time
import pandas as pd

# Logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

def initialize_nbio_com():
    """Initialize the COM object for NBioBSP."""
    try:
        logging.info("Initializing NBioBSP COM...")
        NBioBSP = comtypes.client.CreateObject("NBioBSPCOM.NBioBSP")
        logging.info("NBioBSP COM initialized successfully.")
        return NBioBSP
    except Exception as e:
        logging.error(f"Error initializing NBioBSP COM: {e}")
        return None

def initialize_nbio_device(NBioBSP):
    """Initialize the COM object for NBioBSP.Device."""
    try:
        logging.info("Initializing NBioBSP.Device COM...")
        NBioBSP_Device = NBioBSP.Device
        logging.info("NBioBSP.Device COM initialized successfully.")
        return NBioBSP_Device
    except Exception as e:
        logging.error(f"Error initializing NBioBSP.Device COM: {e}")
        return None

def initialize_nbio_index_search(NBioBSP):
    """Initialize the COM object for NBioBSP.IndexSearch."""
    try:
        logging.info("Initializing NBioBSP.IndexSearch COM...")
        NBioBSP_IndexSearch = NBioBSP.IndexSearch
        logging.info("NBioBSP.IndexSearch COM initialized successfully.")
        return NBioBSP_IndexSearch
    except Exception as e:
        logging.error(f"Error initializing NBioBSP.IndexSearch COM: {e}")
        return None

def initialize_nbio_extraction(NBioBSP):
    """Initialize the COM object for NBioBSP.Extraction."""
    try:
        logging.info("Initializing NBioBSP.Extraction COM...")
        NBioBSP_Extraction = NBioBSP.Extraction
        logging.info("NBioBSP.Extraction COM initialized successfully.")
        return NBioBSP_Extraction
    except Exception as e:
        logging.error(f"Error initializing NBioBSP.Extraction COM: {e}")
        return None

def main():
    # Global variables
    NBioAPIERROR_NONE = 0
    NBioAPI_FIR_PURPOSE_VERIFY = 1

    # Constant for DeviceID
    NBioAPI_DEVICE_ID_AUTO_DETECT = 255
    
    # Initialize the NBioBSP COM
    m_NBio_Bsp = initialize_nbio_com()
    if m_NBio_Bsp is None:
        return

    m_Device = initialize_nbio_device(m_NBio_Bsp)
    if m_Device is None:
        return

    m_Extraction = initialize_nbio_extraction(m_NBio_Bsp)
    if m_Extraction is None:
        return
    
    m_IndexSearch = initialize_nbio_index_search(m_NBio_Bsp)
    if m_IndexSearch is None:
        return
    
    # Insert the User ID and User Name
    input_user_id = input("Enter the User ID for Registration or leave blank to go directly to Identification: ")
    input_user_name = input("Enter the User Name for Registration: ") if input_user_id else None
    if input_user_id and input_user_name:
        # Check if the user ID already exists in fir.csv
        try:
            df = pd.read_csv('fir.csv', names=['UserID', 'UserName', 'FIR', 'Timestamp'], header=0)
            if not df.empty and input_user_id in df['UserID'].astype(str).values:
                print("Error: User ID already exists.")
                return
        except FileNotFoundError:
            # If the file does not exist, continue normally
            df = pd.DataFrame(columns=['UserID', 'UserName', 'FIR', 'Timestamp'])
            df.to_csv('fir.csv', index=False)  # Create the file with headers if it doesn't exist
                
        # Open the fingerprint device
        m_Device.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
        if m_Device.ErrorCode != NBioAPIERROR_NONE:
            logging.error(f"Error opening fingerprint device: {m_Device.ErrorCode}")
            logging.error(f"Error description: {m_Device.ErrorDescription}")
            print(f"Error opening fingerprint device: {m_Device.ErrorCode}\n{m_Device.ErrorDescription}")
            return
        
        # Capture the fingerprint
        m_Extraction.Enroll(input_user_id, 0)
        if m_Extraction.ErrorCode != NBioAPIERROR_NONE:
            logging.error(f"Error capturing fingerprint: {m_Extraction.ErrorCode}")
            logging.error(f"Error description: {m_Extraction.ErrorDescription}")
            print(f"Error capturing fingerprint: {m_Extraction.ErrorCode}\n{m_Extraction.ErrorDescription}")
            return
        
        # Close the fingerprint device
        m_Device.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
        if m_Device.ErrorCode != NBioAPIERROR_NONE:
            logging.error(f"Error closing fingerprint device: {m_Device.ErrorCode}")
            logging.error(f"Error description: {m_Device.ErrorDescription}")
            print(f"Error closing fingerprint device: {m_Device.ErrorCode}\n{m_Device.ErrorDescription}")
            return
        
        # Save the fingerprint to fir.csv
        with open('fir.csv', 'a') as file:
            file.write(f"{input_user_id},{input_user_name},{m_Extraction.TextEncodeFIR},{time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        logging.info("Fingerprint saved successfully.")
        print("Success: Fingerprint saved successfully.")
    
    # Load the content of fir.csv into a table using pandas
    try:        
        df = pd.read_csv('fir.csv', names=['UserID', 'UserName', 'FIR', 'Timestamp'])
        logging.info("Content of fir.csv loaded successfully.")
        
        # Clear the IndexSearch
        m_IndexSearch.ClearDB()
        
        # Add the fingerprints to IndexSearch
        for _, row in df.iterrows():
            m_IndexSearch.AddFIR(row['FIR'], row['UserID'])
        
        # Display the table
        print(df)
    except Exception as e:
        logging.error(f"Error loading content of fir.csv: {e}")
        print(f"Error loading content of fir.csv: {e}")
        
    # Open the fingerprint device
    m_Device.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
    if m_Device.ErrorCode != NBioAPIERROR_NONE:
        logging.error(f"Error opening fingerprint device: {m_Device.ErrorCode}")
        logging.error(f"Error description: {m_Device.ErrorDescription}")
        print(f"Error opening fingerprint device: {m_Device.ErrorCode}\n{m_Device.ErrorDescription}")
        return    
    
    # Capture the fingerprint for verification
    m_Extraction.Capture(NBioAPI_FIR_PURPOSE_VERIFY)
    if m_Extraction.ErrorCode != NBioAPIERROR_NONE:
        logging.error(f"Error capturing fingerprint for verification: {m_Extraction.ErrorCode}")
        logging.error(f"Error description: {m_Extraction.ErrorDescription}")
        print(f"Error capturing fingerprint for verification: {m_Extraction.ErrorCode}\n{m_Extraction.ErrorDescription}")
        return
    
    # Close the fingerprint device
    m_Device.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
    if m_Device.ErrorCode != NBioAPIERROR_NONE:
        logging.error(f"Error closing fingerprint device: {m_Device.ErrorCode}")
        logging.error(f"Error description: {m_Device.ErrorDescription}")
        print(f"Error closing fingerprint device: {m_Device.ErrorCode}\n{m_Device.ErrorDescription}")
    
    # Encode the captured fingerprint
    TextEncodeFIR = m_Extraction.TextEncodeFIR
    
    # Identify the fingerprint
    m_IndexSearch.IdentifyUser(TextEncodeFIR, 5)
    if m_IndexSearch.ErrorCode != NBioAPIERROR_NONE:
        logging.error(f"Error identifying fingerprint: {m_IndexSearch.ErrorCode}")
        logging.error(f"Error description: {m_IndexSearch.ErrorDescription}")
        print(f"Error identifying fingerprint: {m_IndexSearch.ErrorCode}\n{m_IndexSearch.ErrorDescription}")
        return
    
    # Show the identification result
    if m_IndexSearch.UserID != 0:
        user_id = m_IndexSearch.UserID
        user_name = df.loc[df['UserID'].astype(str) == str(user_id), 'UserName'].values[0]
        print(f"Success: User identified: {user_name} (ID: {user_id})")
    else:
        print(f"Failure: User not identified: {m_IndexSearch.ErrorCode}")

if __name__ == "__main__":
    main()