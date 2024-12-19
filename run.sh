#!/bin/bash

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

set -e 

command -v python3 >/dev/null 2>&1 || { echo >&2 "Python3 is required but not installed. Aborting."; exit 1; }
command -v virtualenv >/dev/null 2>&1 || { python3 -m pip install --user virtualenv; }

gitignore() {
    echo -e "${YELLOW}♠︎ Generating .gitignore file${NC}"
    cat > .gitignore << EOL
.vscode
__pycache__
*.pyc
.venv
.env
EOL
}

createStructure(){
    mkdir -p inputs/{excel,word,pdf} outputs/{reports,processed} tests resources
    touch tests/file_analysis.robot
    touch resources/excel_analyzer.py
    touch resources/word_analyzer.py
    touch resources/pdf_analyzer.py
    touch resources/report_generator.py
}

createExcelAnalyzer() {
    echo -e "${YELLOW}🚀 Creating Excel Analyzer${NC}"
    cat > resources/excel_analyzer.py << EOL
import pandas as pd
import openpyxl

def analyze_excel_file(file_path):
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        
        # Basic analysis
        analysis = {
            'total_rows': len(df),
            'total_columns': len(df.columns),
            'column_names': list(df.columns),
            'numeric_columns': list(df.select_dtypes(include=['number']).columns),
            'summary_statistics': {}
        }
        
        # Generate summary statistics for numeric columns
        for col in analysis['numeric_columns']:
            analysis['summary_statistics'][col] = {
                'mean': df[col].mean(),
                'median': df[col].median(),
                'min': df[col].min(),
                'max': df[col].max(),
                'std_dev': df[col].std()
            }
        
        # Check for missing values
        analysis['missing_values'] = df.isnull().sum().to_dict()
        
        return analysis
    except Exception as e:
        return {'error': str(e)}
EOL
}

createWordAnalyzer() {
    echo -e "${YELLOW}🚀 Creating Word Analyzer${NC}"
    cat > resources/word_analyzer.py << EOL
from docx import Document

def analyze_word_file(file_path):
    """
    Comprehensive Word document analysis
    """
    try:
        doc = Document(file_path)
        
        analysis = {
            'total_paragraphs': len(doc.paragraphs),
            'total_tables': len(doc.tables),
            'total_images': len([shape for paragraph in doc.paragraphs for shape in paragraph._element.findall('.//a:graphic', namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})]),
            'paragraphs': [],
            'word_count': 0
        }
        
        # Analyze paragraphs
        for para in doc.paragraphs:
            para_analysis = {
                'text': para.text,
                'words': len(para.text.split()),
                'style': para.style.name if para.style else 'No Style'
            }
            analysis['paragraphs'].append(para_analysis)
            analysis['word_count'] += len(para.text.split())
        
        return analysis
    except Exception as e:
        return {'error': str(e)}
EOL
}

createPDFAnalyzer() {
    echo -e "${YELLOW}🚀 Creating PDF Analyzer${NC}"
    cat > resources/pdf_analyzer.py << EOL
import PyPDF2

def analyze_pdf_file(file_path):
    """
    Comprehensive PDF file analysis
    """
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            analysis = {
                'total_pages': len(reader.pages),
                'is_encrypted': reader.is_encrypted,
                'page_details': []
            }
            
            # Analyze each page
            for page_num, page in enumerate(reader.pages):
                page_analysis = {
                    'page_number': page_num + 1,
                    'text': page.extract_text(),
                    'word_count': len(page.extract_text().split()),
                }
                analysis['page_details'].append(page_analysis)
            
        return analysis
    except Exception as e:
        return {'error': str(e)}
EOL
}

createReportGenerator() {
    echo -e "${YELLOW}🚀 Creating Report Generator${NC}"
    cat > resources/report_generator.py << EOL
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch

def generate_analysis_report(analyses, output_path):
    """
    Generate a comprehensive PDF report
    """
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    
    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawString(inch, height - inch, "Comprehensive File Analysis Report")
    
    # Excel Analysis Section
    c.setFont("Helvetica", 12)
    c.drawString(inch, height - 2*inch, "Excel File Analysis")
    
    # Add more report generation logic here
    c.save()
EOL
}

createRobot() {
    echo -e "${YELLOW}🚀 Creating Robot${NC}"
    cat > tests/file_analysis.robot << EOL
*** Settings ***
Documentation     File Analysis RPA
Library           OperatingSystem
Library           Collections
Library           Process
Resource          ../resources/excel_analyzer.py
Resource          ../resources/word_analyzer.py
Resource          ../resources/pdf_analyzer.py
Resource          ../resources/report_generator.py

*** Variables ***
${input_excel_dir}    /inputs/excel
${input_word_dir}     /inputs/word
${input_pdf_dir}      /inputs/pdf
${output_report_path} /outputs/reports/comprehensive_report.pdf
@{analysis_results}    # List to store analysis results


*** Test Cases ***
Analyze Excel Files
    @{excel_files}=    List Files In Directory    ${input_excel_dir}    *.xlsx
    FOR    ${file}    IN    @{excel_files}
        ${excel_analysis}=    Analyze Excel File    ${file}  # Assuming Analyze Excel File returns a dictionary
        Append To List    ${analysis_results}    ${excel_analysis}
        Log Dictionary    ${excel_analysis}
    END

Analyze Word Documents
    @{word_files}=    List Files In Directory    ${input_word_dir}    *.docx
    FOR    ${file}    IN    @{word_files}
        ${word_analysis}=    Analyze Word File    ${file} # Assuming Analyze Word File returns a dictionary
        Append To List    ${analysis_results}    ${word_analysis}
        Log Dictionary    ${word_analysis}
    END

Analyze PDF Files
    @{pdf_files}=    List Files In Directory    ${input_pdf_dir}    *.pdf
    FOR    ${file}    IN    @{pdf_files}
        ${pdf_analysis}=    Analyze PDF File    ${file} # Assuming Analyze PDF File returns a dictionary
        Append To List    ${analysis_results}    ${pdf_analysis}
        Log Dictionary    ${pdf_analysis}
    END


Generate Comprehensive Report
    Generate Analysis Report    ${output_report_path}    ${analysis_results}  # Pass the collected results
EOL
}

createScripts(){
    createExcelAnalyzer
    createPDFAnalyzer
    createReportGenerator
    createWordAnalyzer
    createRobot
}

main() {
    echo -e "${YELLOW}🔧 Audio Recognition Application Initialization${NC}"

    touch .gitignore .env
    gitignore
    createStructure
    createScripts

    python3 -m venv .venv
    source .venv/bin/activate
    pip install --upgrade pip setuptools wheel
    pip install --upgrade pip
    pip install robotframework robotframework-pythonlibcore openpyxl python-docx pypdf2 pandas reportlab docx

    echo -e "${GREEN}🎉 Project is ready! run 'robot tests/file_analysis.robot' to start.${NC}"
}

main

