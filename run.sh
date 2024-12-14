GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

set -e

command -v python3 >/dev/null 2>&1 || { echo >&2 "Python3 is required but not installed.  Aborting."; exit 1; }
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
    mkdir -p rpa/inputs/{excel,word,pdf} rpa/outputs/{reports,processed} rpa/tests rpa/resources

    touch rpa/tests/file_analysis.robot
    touch rpa/resources/excel_analyzer.py
    touch rpa/resources/word_analyzer.py
    touch rpa/resources/pdf_analyzer.py
    touch rpa/resources/report_generator.py
}

main() {
    echo -e "${YELLOW}🔧 Audio Recognition Application Initialization${NC}"

    touch .gitignore .env
    gitignore

    python3 -m venv .venv
    source .venv/bin/activate
    pip install --upgrade pip setuptools wheel
    pip install --upgrade pip
    pip install robotframework robotframework-pythonlibcore openpyxcl python-docx pypdf2 pandas reportlab

    echo -e "${GREEN}🎉 Project is ready! run 'robot tests/file_analysis.robot' to start.${NC}"
}

main

