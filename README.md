# AutoLabels

This is intended to be used by people at my company. One button press will perform many actions to result in beautifully created labels and reports.

1. The program starts by acquiring a bearer token to make a REST API request. It does this by navigating the webbrowser and extracting cookies (unable to make request to get bearer token directly from MFA due to privileges)
2. Send an email to run a Power Automate script to extract information from microsoft planner.
3. Power automate sends an email back with the information and the information is then extracted into separate blocks. 
   Information such as a part number, storage vessels, etc.
5. Each block has a corresponding number of variables that need to be determined. 
6. API requests are then made for each block to look up the required information. The program converts a doc or docx to text and reads through it.
7. SAP is then opened and navigated to find the last bit of information required.
8. Labels are created in excel with proper formatting and page breaks.
9. Ignition is opened to create reports for a multitude of possible scenarios with failsafes and verification steps to ensure the report is correct.

Note: Some elements in the code have been removed for privacy (mainly file paths and company sensitive information).
