import comtypes.client
import logging
import time
import tkinter as tk
from tkinter import simpledialog, messagebox
import pandas as pd
from pandastable import Table

# Logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

# Global variables
NBioAPIERROR_NONE = 0
NBioAPI_FIR_PURPOSE_VERIFY = 1
NBioAPI_DEVICE_ID_AUTO_DETECT = 255

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

def register_user(m_Device, m_Extraction):
    """Function to register a new user."""
    input_user_id = simpledialog.askstring("User ID Input", "Enter the User ID for Registration:")
    input_user_name = simpledialog.askstring("User Name Input", "Enter the User Name for Registration:")
    if input_user_id and input_user_name:
        try:
            df = pd.read_csv('fir.csv', names=['UserID', 'UserName', 'FIR', 'Timestamp'], header=0)
            if not df.empty and input_user_id in df['UserID'].astype(str).values:
                messagebox.showerror("Error", "User ID already exists.")
                return
        except FileNotFoundError:
            df = pd.DataFrame(columns=['UserID', 'UserName', 'FIR', 'Timestamp'])
        
        m_Device.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
        if m_Device.ErrorCode != NBioAPIERROR_NONE:
            logging.error(f"Error opening fingerprint device: {m_Device.ErrorCode}")
            messagebox.showerror("Error", f"Error opening fingerprint device: {m_Device.ErrorCode}")
            return
        
        m_Extraction.Enroll(input_user_id, 0)
        if m_Extraction.ErrorCode != NBioAPIERROR_NONE:
            logging.error(f"Error capturing fingerprint: {m_Extraction.ErrorCode}")
            messagebox.showerror("Error", f"Error capturing fingerprint: {m_Extraction.ErrorCode}")
            return
        
        m_Device.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
        if m_Device.ErrorCode != NBioAPIERROR_NONE:
            logging.error(f"Error closing fingerprint device: {m_Device.ErrorCode}")
            messagebox.showerror("Error", f"Error closing fingerprint device: {m_Device.ErrorCode}")
            return
        
        with open('fir.csv', 'a') as file:
            file.write(f"{input_user_id},{input_user_name},{m_Extraction.TextEncodeFIR},{time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        logging.info("Fingerprint saved successfully.")
        messagebox.showinfo("Success", "Fingerprint saved successfully.")

def view_registered_users():
    """Function to view the list of registered users."""
    try:
        df = pd.read_csv('fir.csv', names=['UserID', 'UserName', 'FIR', 'Timestamp'])
        logging.info("Content of fir.csv loaded successfully.")
        
        root = tk.Toplevel()
        root.title("List of Registered Users")
        frame = tk.Frame(root)
        frame.pack(fill='both', expand=True)
        
        table = Table(frame, dataframe=df)
        table.show()
        table.redraw()
        
        root.mainloop()
    except Exception as e:
        logging.error(f"Error loading content of fir.csv: {e}")
        messagebox.showerror("Error", f"Error loading content of fir.csv: {e}")

def identify_user(m_Device, m_Extraction, m_IndexSearch):
    """Function to identify a user."""
    
    df = pd.read_csv('fir.csv', names=['UserID', 'UserName', 'FIR', 'Timestamp'])
    logging.info("Content of fir.csv loaded successfully.")
    
    # Clear the IndexSearch
    m_IndexSearch.ClearDB()
    
    # Add fingerprints to IndexSearch
    for _, row in df.iterrows():
        m_IndexSearch.AddFIR(row['FIR'], row['UserID'])
    
    m_Device.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
    if m_Device.ErrorCode != NBioAPIERROR_NONE:
        logging.error(f"Error opening fingerprint device: {m_Device.ErrorCode}")
        messagebox.showerror("Error", f"Error opening fingerprint device: {m_Device.ErrorCode}")
        return
    
    m_Extraction.Capture(NBioAPI_FIR_PURPOSE_VERIFY)
    if m_Extraction.ErrorCode != NBioAPIERROR_NONE:
        logging.error(f"Error capturing fingerprint for verification: {m_Extraction.ErrorCode}")
        messagebox.showerror("Error", f"Error capturing fingerprint for verification: {m_Extraction.ErrorCode}")
        return
    
    m_Device.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
    if m_Device.ErrorCode != NBioAPIERROR_NONE:
        logging.error(f"Error closing fingerprint device: {m_Device.ErrorCode}")
        messagebox.showerror("Error", f"Error closing fingerprint device: {m_Device.ErrorCode}")
    
    TextEncodeFIR = m_Extraction.TextEncodeFIR
    m_IndexSearch.IdentifyUser(TextEncodeFIR, 5)
    if m_IndexSearch.ErrorCode != NBioAPIERROR_NONE:
        logging.error(f"Error identifying fingerprint: {m_IndexSearch.ErrorCode}")
        messagebox.showerror("Error", f"Error identifying fingerprint: {m_IndexSearch.ErrorCode}")
        return
    
    if m_IndexSearch.UserID != 0:
        user_id = m_IndexSearch.UserID
        user_name = df.loc[df['UserID'].astype(str) == str(user_id), 'UserName'].values[0]
        messagebox.showinfo("Success", f"User identified: {user_name} (ID: {user_id})")
    else:
        messagebox.showinfo("Failure", "User not identified")

def main():
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
    
    # Create the graphical interface
    root = tk.Tk()
    root.title("Fingerprint System")
    
    # Set window size
    window_width = 400
    window_height = 300

    # Get the screen dimension
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Find the center point
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)

    # Set the position of the window to the center of the screen
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    # Create buttons for each option
    btn_register = tk.Button(root, text="Register User", command=lambda: register_user(m_Device, m_Extraction))
    btn_register.pack(pady=10)

    btn_view = tk.Button(root, text="View Registered Users", command=view_registered_users)
    btn_view.pack(pady=10)

    btn_identify = tk.Button(root, text="Identify User", command=lambda: identify_user(m_Device, m_Extraction, m_IndexSearch))
    btn_identify.pack(pady=10)

    btn_exit = tk.Button(root, text="Exit", command=root.quit)
    btn_exit.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    main()