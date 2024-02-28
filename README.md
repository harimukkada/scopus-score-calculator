# Scopus Publication Data Exporter
![GitHub Logo](https://www.elsevier.com/images/elsevier-logo.svg)



The Scopus Publication Data Exporter is a Flask-based web application designed to simplify the extraction and exportation of publication data associated with Scopus author IDs. This tool leverages the Pybliometrics library to access the Scopus database and retrieve comprehensive publication details, enabling researchers and academics to efficiently organize and analyze their scholarly output.

## Features
**Web Interface:** Users can input Scopus author IDs via a user-friendly HTML form, making the process intuitive and accessible.  
**Data Retrieval:** The application fetches publication information from Scopus using author IDs, ensuring accurate and up-to-date data retrieval.  
**Data Processing:** Various data processing tasks are performed, including author information extraction, citation count formatting, and calculation of citation points and quartiles.  
**Excel Export:** Processed publication data is exported to an Excel file format, facilitating further analysis and sharing of research findings.  
**Additional Functionalities:** Advanced features such as percentile calculation, mark assignment based on percentile or citation points, and computation of total points are provided to enhance data analysis capabilities.  
**Error Handling:** The application includes robust error handling mechanisms to manage scenarios where no results are found for the provided author ID, ensuring a smooth user experience.  

## How to Use
&bull; Clone this repository to your local machine.  
&bull; Install the required dependencies by running pip install -r requirements.txt.  
&bull; Run the Flask application by executing python app.py.  
&bull; Access the application through your web browser at http://localhost:5000.  
&bull; Input the desired Scopus author ID in the provided form and submit.  
&bull; The application will retrieve the publication data associated with the provided author ID and generate an Excel file for download.  

## Contributing
Contributions to the Scopus Publication Data Exporter are welcome! Whether it's bug fixes, feature enhancements, or documentation improvements, feel free to submit a pull request.

## License:
This project is licensed under the MIT License.

## Acknowledgments
Special thanks to the developers of Flask, Pybliometrics, and pandas for their excellent libraries and tools that made this project possible.

## Contact
For any inquiries or feedback, please contact hariavailablehere@gmail.com.com.

Enjoy exploring and analyzing your publication data with the Scopus Publication Data Exporter!
