# Bundle Slip Mail Automator

Bundle Slip Mail Automator is a Python script designed to automate the process of fetching emails from a specific email account, extracting attachments, sorting and arranging them, and saving them into a designated folder.

## Features

- Fetches emails from a specified email account using the Gmail API.
- Extracts attachments (such as PDFs, images, and Excel files) from the emails.
- Performs various operations on the attachments, including:
  - Cropping images to the desired size.
  - Converting images to PDF format.
  - Merging multiple PDF files into a single file.
  - (Optionally) Merging multiple Excel files into a single file.
- Moves processed files to respective folders for better organization.
- Provides an additional function to extract email addresses from `.eml` files and store them in a JSON file.

## Prerequisites

Before running the script, make sure you have the following prerequisites:

- Python 3.x installed on your system.
- A Google account with Gmail enabled.
- A `credentials.json` file obtained from the Google Cloud Console (instructions provided in the project files).

## Installation

1. Clone the repository or download the source code:

   ```bash
   git clone https://github.com/abhi245y/bundle-slip-mail-automator.git
   ```

2. Navigate to the project directory:

   ```bash
   cd bundle-slip-mail-automator
   ```

3. Run the `install_requirements.sh` script to install the required Python packages:

   ```bash
   ./install_requirements.sh
   ```

## Configuration

1. Obtain the `credentials.json` file from the Google Cloud Console by following these steps:
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a new project or select an existing one.
   - Enable the Gmail API for your project.
   - Create credentials (an OAuth client ID) for a desktop application.
   - Download the `credentials.json` file and place it in the project directory.

2. (Optional) If your college/organization email addresses are not included in the `config/college_emails.json` file, you can update it manually or use the `extraFunction()` in `main.py` to extract email addresses from `.eml` files and store them in the JSON file.

## Usage

1. Run the `runScript.sh` script to start the script:

   ```bash
   ./runScript.sh
   ```

2. The script will prompt you to log in to your Google account if necessary.
3. After successful authentication, the script will fetch emails from the specified account, process the attachments, and save the output files in the designated folders.

## Project Structure

```
bundle-slip-mail-automator/
├── config/
│   └── college_emails.json
├── EZGmail.py
├── install_requirements.sh
├── main.py
├── runScript.sh
├── runScript.bat
└── README.md
```

- `main.py`: The main Python script that contains the core functionality.
- `EZGmail.py`: A Python module used to interact with the Gmail API (created by Al Sweigart).
- `config/college_emails.json`: A JSON file containing the list of college/organization email addresses used for filtering emails.
- `install_requirements.sh`: A bash script to install the required Python packages.
- `runScript.sh`: A bash script to run the `main.py` script.
- `README.md`: This file, providing instructions and documentation for the project.

## Contributing

Contributions to this project are welcome. If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## Credits

- The `EZGmail.py` module was created by Al Sweigart (al@inventwithpython.com).
- The project utilizes various Python libraries, including `ezgmail`, `img2pdf`, `PyPDF2`, `imutils`, `cv2`, and `pandas`.
