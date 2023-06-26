Web Phone Number Extractor (WPNE) 1.2
WPNE is a Python application that allows you to extract phone numbers from websites. It sends GET requests to the specified websites, parses the HTML content, and extracts phone numbers matching various formats. You can also specify a delay between each request to simulate human-like behavior.

Features
Extracts phone numbers from websites
Supports multiple phone number formats
Filters out toll-free numbers
Provides a progress bar to track the processing
Pauses and resumes the processing
Saves the results to an Excel file
Requirements
Python 3.x
tkinter module (included in the standard library)
openpyxl library (pip install openpyxl)
pandas library (pip install pandas)
requests library (pip install requests)
beautifulsoup4 library (pip install beautifulsoup4)
Usage
Install the required dependencies listed above.
Run the wpne.py script using Python: python wpne.py.
The application window will open.
Click the "Start Application" button to begin.
Enter the delay in seconds between each request (1 to n) in the entry field provided.
Click the "Upload Websites File" button and select an Excel file (.xlsx) containing the list of websites to process. The file should have a sheet named "Sheet1" with a column named "Website" containing the website URLs.
The processing will begin, and the progress bar will show the progress. The website being processed and the fetched phone number(s) will be displayed.
You can pause the processing by clicking the "Pause" button. To resume, click the "Resume" button.
Once the processing is complete, a dialog box will appear prompting you to save the results file. Choose a location and provide a name for the file. The results will be saved in an Excel file (.xlsx) format.
Click the "Download Results File" button to download the saved results file.
Note: The application uses a randomized delay between each request to simulate human-like behavior and avoid overwhelming the websites. The delay value is specified by the user.

Important Notes
The websites file should be in Excel format (.xlsx) and should contain a sheet named "Sheet1" with a column named "Website" containing the website URLs.
Ensure that the websites in the file are valid and accessible.
The application may take some time to process a large number of websites or websites with slow response times.
The application relies on the accuracy and consistency of the phone number patterns. It may not capture all phone numbers in certain cases or with complex formatting.
License
This project is licensed under the MIT License.

Feel free to modify and enhance the code according to your needs.

Acknowledgments
The WPNE application was developed by Peaskeepper.
It utilizes the OpenPyXL library for working with Excel files.
It also uses the Pandas library for data manipulation and analysis.
The Requests library is used for sending HTTP requests.
The Beautiful Soup library is used for parsing HTML content.
Contributing
Contributions to the project are welcome! If you encounter any issues, have suggestions for improvements, or would like to add new features, please submit an issue or a pull request.

