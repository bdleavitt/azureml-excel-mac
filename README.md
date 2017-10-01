# azureml-excel-mac
Azure Machine Learning Studio is a neat piece of web-based software that lets you create and deploy machine learning models without a lot of coding. You can consume your deployed models as a web service, making it easy to incorporate machine learning into applications. You can even run them from Excel... BUT ONLY IF YOU'RE ON A PC (as of this writing, anyway - 9/30/2017). This project aims to provide a Mac-compatible VBA project to allow our Excel-for-Mac friends to consume Azure Machine Learning web services like the rest of us. It should also work on your Windows-based PC. 

## Getting Started
1. To get started, just download a copy of the Web Enabled Workbook. 
2. You will need to enable Macros. 
3. On the Home Tab, you'll see a button "Get Predictions from Web". Click on it.
4. From the Azure Machine Learning website, you will need to find the URL and API Key. (Note: just provide the base URL, which is everything up to and including the last slash "/". 
5. Choose where to get your data from, or click "Load Sample Data" (if you haven't specified the input range, it will go to whatever cell is highlighted).
6. Choose where to put the result of your data (again... if not specified, the results will be output to wherever the selected cell is.
7. Click "Predict" - if everything goes according to plan, you should be rewarded with a fresh bunch of predictions from the API. 

## Acknowledgments
This project relies on VBA-WEB, a project from Tim Hall and friends over at VBA-TOOLS: https://github.com/VBA-tools/VBA-Web/, which is distributed under an MIT License. This project is also made available under an MIT License. 

