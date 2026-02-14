 Retail Analysis - Honda Campaign

Automates the processing of call and phone files by cross-referencing calls, phones, and occurrences, generating detailed Excel reports.

 Description

This project processes Honda campaign data, performing:

Reading the most recent Excel files for calls and occurrences.

Reading CSV phone files with multiple layouts.

Normalizing phone numbers and identifying records by CPF.

Cross-referencing calls, phones, and "De X Para" occurrences.

Generating reports:

DETAIL: all calls matched with phones.

BY_PHONE: summary per phone.

PIVOT_SUMMARY: service-level summary with occurrence counts.

The script can process multiple CSVs automatically, handling different phone layouts.

Requirements

Python >= 3.8

Python libraries:

pip install pandas openpyxl


Access to network folders containing input files.

Folder Structure
File Type	Expected Path
Calls (Excel)	PASTA_CHAMADAS
Occurrences (Excel)	PASTA_OCORRENCIAS
Phones (CSV)	PASTA_TELEFONES
Output Files	PASTA_SAIDA

All paths can be configured directly in the Python script.

How to Use

Clone the repository:

git clone https://github.com/Cardoso-leo/honda-analysis.git
cd honda-analysis


Install dependencies:

pip install pandas openpyxl


Configure input and output folders in the script (PASTA_CHAMADAS, PASTA_OCORRENCIAS, etc.)

Run the script:

python analise_honda.py


Output files will be saved in the output folder with the prefix ANALISE_.

Output

Each generated Excel file contains three sheets:

DETAIL – all calls matched with phones.

BY_PHONE – summary of calls per phone number.

PIVOT_SUMMARY – service-level summary showing counts of alo, cpc, and promessa.

 Notes

CPF is used as the main identifier.

Phone numbers are normalized (removing spaces, symbols, and country code 55).

If a phone CSV is not recognized, the script continues with the next file.

Supported layouts:

ddd01 + telefone01, ddd02 + telefone02 …

ddd + numero

Possible Improvements

Automatically detect CPF and phone columns for more robustness.

Generate reports in PDF or interactive dashboards.

Add detailed logs showing execution time per file.
