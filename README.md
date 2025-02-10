# eNBioBSPcom.PY

## Overview
This project provides a graphical and command-line interface for fingerprint registration and identification using the NBioBSP COM library.

## Prerequisites
- [Python 3.x](https://www.python.org/downloads/)
- [pip (Python package installer)](https://packaging.python.org/en/latest/tutorials/installing-packages/)
- [Interop.NBioBSPCOMLib.dll](https://suporte.fingertech.com.br/portal-do-desenvolvedor/)

## Clone the Repository
To clone the repository, run the following command:
```sh
git clone https://github.com/yourusername/eNBioBSPcom.PY.git
cd eNBioBSPcom.PY
```

## Install Dependencies
To install the required dependencies, run:
```bash
pip install -r requirements.txt
```

## Usage

### Command-Line Interface
To use the command-line interface, run:
```bash
python IndexSearch.py
```

### Graphical Interface
To use the graphical interface, run:
```bash
python IndexSearch_graph.py
```

## List Methods
To see the list of methods available in `nbio_bsp`, run:
```python
for n in dir(nbio_bsp):
        print(n)
```
To see the list of methods available in `nbio_bsp.IndexSearch`, run:
```python
for n in dir(nbio_bsp.IndexSearch):
        print(n)
```

## Features
- Register User: Register a new user with their fingerprint.
- View Registered Users: View the list of registered users.
- Identify User: Identify a user based on their fingerprint.

## File Descriptions
- IndexSearch.py: Contains the command-line interface for fingerprint registration and identification.
- IndexSearch_graph.py: Contains the graphical interface for fingerprint registration and identification.
- fir.csv: Stores the registered users' information and fingerprints.
- requirements.txt: Lists the required Python packages.

## Logging
Logs are generated to provide information about the operations and any errors encountered. The logs are displayed in the console.

## Error Handling
Error messages are logged and displayed to the user in case of any issues during the operations.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
