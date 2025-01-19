# Email Supply Chain Report Saver

## Overview

The **Email Supply Chain Report Saver** script is an automation tool designed to download attachments from Outlook emails automatically. It is specifically built to help supply chain professionals, businesses, and ERP systems collect reports sent via email and save them locally. This helps streamline and digitize the supply chain process, reduce manual efforts, and align with Industry 4.0 principles.

This tool interacts with Microsoft Outlook, searches for emails based on specific subject lines, and downloads their attachments (such as Excel files, PDFs, or CSVs) to a designated directory on the local machine.

---

## Features

- **Automates the process of saving email attachments**: Automatically fetches attachments from Outlook emails, eliminating the need for manual downloads.
- **Customizable folder and subject line search**: Allows you to specify the folder in Outlook and the subject of the email for the report.
- **Organizes saved files**: Saves files in a structured directory to ensure that all reports are stored in one central location.
- **Scalable and extendable**: Easily adaptable for different supply chain departments or email accounts.

---

## Prerequisites

To use this script, the following dependencies and tools need to be installed and set up:

1. **Python 3.x**: The script is written in Python 3.x. Ensure you have Python installed on your system.
   - Download from [Python.org](https://www.python.org/downloads/)
   
2. **Microsoft Outlook**: This script is designed to work with Microsoft Outlook. You must have Outlook installed and configured on your machine.
   
3. **pywin32 Library**: The `pywin32` library is required to interface with Microsoft Outlook using Python. It allows Python to interact with COM objects, which is essential for accessing Outlook.
   - To install, run:
     ```
     pip install pypiwin32
     ```

---

## Installation

### Step 1: Install Python

If you don't already have Python installed, download and install it from the official website: [https://www.python.org/downloads/](https://www.python.org/downloads/).

### Step 2: Install Required Libraries

Install the `pywin32` library, which is necessary for interacting with Outlook:

```bash
pip install pypiwin32
```

### Step 3: Ensure Microsoft Outlook is Installed

Make sure that Microsoft Outlook is installed and configured on your system. The script relies on Outlook's ability to access emails and download attachments.

---

## Usage

### Step 1: Download the Script

Clone this repository to your local machine or download the script file:

```bash
git clone https://github.com/yourusername/email-supply-chain-report-saver.git
```

### Step 2: Configure the Script

Before running the script, you need to modify a few parameters:

1. **Select Outlook Account**:
   - In the script, the connection to Outlook is created using `outlook.Folders.Item(2)`. The number `2` represents the second account in the list. Change this number to `1` if you want to access your primary account.
   
   ```python
   root_folder = outlook.Folders.Item(2)  # Change 2 to 1 if you want the primary account
   ```

2. **Specify the Folder Name**:
   - Replace `'Your_Folder_Name'` with the folder in which the reports are stored (e.g., "Inbox", "Archive", or any custom folder).
   
   ```python
   target_folder_name = 'Your_Folder_Name'
   ```

3. **Customize the Email Subject and File Name**:
   - In the `save_attachments()` function call, replace `"Your Email Subject Here"` with the exact subject of the email you want to search for. Modify the `file_name` parameter to specify how you want to save the attachments.
   
   ```python
   save_attachments(
       subject="Your Email Subject Here",  # Replace with the email subject
       file_name="Downloaded_Attachment.xlsx"  # Replace with the desired filename
   )
   ```

### Step 3: Run the Script

Once the configuration is done, simply run the script. It will automatically connect to your Outlook, search for emails with the specified subject, and download the attachments to your local machine.

```bash
python email_supply_chain_report_saver.py
```

The script will:
- Search the specified folder for emails with the matching subject.
- Download any attachments from the emails and save them in the `Outlook_Attachments` folder within your `Downloads` directory.

---

## Example Output

Upon successful execution, the script will print a summary of its actions in the console. If the email with the specified subject is found, you will see messages like this:

```
Working with email account: <Your Outlook Email Address>
Found email with subject: Supply Chain Report
Saved attachment as: C:\Users\<YourUsername>\Downloads\Outlook_Attachments\Downloaded_Attachment.xlsx
```

If no matching email is found, the script will output:

```
No email found with subject: Supply Chain Report
```

---

## Customization

The script can be customized to fit different use cases or scale up for use with multiple reports, different file types, or various subject lines. Here are some suggestions:

- **Save multiple attachments**: By default, the script saves only the first attachment found. You can remove the `break` statement inside the loop if you want to save all attachments from a matching email.
  
  ```python
  for attachment in message.Attachments:
      attachment.SaveAsFile(save_path)
  ```

- **Automate the process**: You can schedule this script to run at regular intervals using tools like Windows Task Scheduler (on Windows) or cron jobs (on Linux/macOS).

---

## Benefits

1. **Automation**: The script automates the collection and storage of supply chain reports, reducing manual effort and time.
2. **Digitization**: Helps digitize supply chain processes by organizing reports in a digital format for easy access and analysis.
3. **Efficiency**: Saves significant time, especially when managing large volumes of emails and attachments from multiple suppliers or ERP systems.
4. **Scalability**: Easily extendable to accommodate multiple report types, accounts, and file formats as the business grows.

---

## Contributing

We welcome contributions! If you'd like to improve the script, please fork the repository and submit a pull request. Here's how you can contribute:

1. Fork the repository.
2. Create a new branch.
3. Make your changes and commit them.
4. Push to your forked repository.
5. Submit a pull request with a clear description of your changes.

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Acknowledgments

- **pywin32**: For providing the library that allows interaction with Outlook.
- **Microsoft Outlook**: For providing the email platform to work with.

---

With this automation, you can simplify the management of supply chain reports and free up valuable time to focus on more strategic tasks.
