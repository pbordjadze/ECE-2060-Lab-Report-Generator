from docx import Document
from datetime import datetime, timedelta
import re
import sys
import os

def change_extension(path):
    # Split the path into root and extension
    root, ext = os.path.splitext(path)
    # Replace the extension with ".tex"
    new_path = root + ".tex"
    return new_path

# Check if the correct number of arguments is provided
if len(sys.argv) != 2:
    print("Usage: python latex_generator.py filename.docx")
    sys.exit(1)

# Extract the command-line argument
argument = sys.argv[1]

doc = Document(argument)
lines = [paragraph.text for paragraph in doc.paragraphs]

pattern = r'\b\d+\b'
s = lines.pop(0)

# Find the match in the string
match = re.search(pattern, s)

if match:
    # Extract the integer from the matched group
    lab_number = int(match.group())
else:
    print("No match found")


# Initialize variables
experimental_methodology_items = []
results_items = []
number_of_questions = 0
state = "start"  # Initial state


for string in lines:
    if state == "start":
        if string.startswith("Experimental Methodology"):
            state = "experimental_methodology"
    elif state == "experimental_methodology":
        if string.startswith("Results"):
            state = "results"
        else:
            experimental_methodology_items.append(string)
    elif state == "results":
        if string.startswith("Discussion"):
            state = "discussion"
        else:
            results_items.append(string)
    elif state == "discussion":
        if string.startswith("Additional Comments"):
            break
        else:
            number_of_questions += 1






def snake_case_without_parentheses(input_string):
    # Remove numbers and periods within parentheses
    cleaned_string = re.sub(r'\(\d+(\.\d+)?/\d+(\.\d+)?\)', '', input_string)
    
    # Convert string to snake case
    cleaned_string = cleaned_string.rstrip().replace(" ", "_").lower()
    
    return cleaned_string





DATE_OF_FIRST_LAB = datetime(2024, 2, 2)


lab_title = "Default Title"
lab_date = DATE_OF_FIRST_LAB + timedelta(weeks=(lab_number - 1))





headers = """
\\documentclass{article}
\\usepackage{graphicx}
\\usepackage{enumitem}
\\usepackage{lipsum}
\\usepackage{mdframed}
\\usepackage{amsmath}
\\usepackage{underscore}
\\usepackage{parskip}

\\newcommand{\\includegraphicsconditionally}[1]{%
  \\IfFileExists{#1}{%
    \\includegraphics[width=\\textwidth]{#1}%
  }{%
    \\textit{Placeholder: #1}
  }%
}

\\newmdenv[linecolor=black, linewidth=2pt]{answerbox}

"""

title_string = f"\\title{{Lab Report \\#{lab_number}\\\\{lab_title}}}\n"

authors_string = "\\author{Paul Bordjadze, Suraj Rao \\\\Group 23, 8 AM}\n"

date_string = f"\\date{{{lab_date.strftime('%m/%d/%Y')}}}\n"

begin_document_string = "\\begin{document}\n"

make_title_string = "\\maketitle\n\\newpage"

executive_summary_string = "\\section{Executive Summary}\n"

introduction_string = "\\section{Introduction}\n"

experimental_methodology_string = "\\section{Experimental Methodology}\n"

for item in experimental_methodology_items:
    item = snake_case_without_parentheses(item)
    item_string = f"\\includegraphicsconditionally{{lab{lab_number}_images/{item}.png}}\\\\\n"
    experimental_methodology_string = experimental_methodology_string + item_string

results_string = "\\section{Results}\n"

for item in results_items:
    item = snake_case_without_parentheses(item)
    item_string = f"\\includegraphicsconditionally{{lab{lab_number}_images/{item}.png}}\\\\\n\n"
    results_string = results_string + item_string

discussion_string = """\section{Discussion}\n\\begin{enumerate}[label=(\\arabic*), left=0pt, labelsep=12pt, itemsep=1em]\n"""

question_string = """
        \\item QUESTION
        \\begin{answerbox}
        ANSWER
        \\end{answerbox}"""
discussion_string = discussion_string + (question_string * number_of_questions) + "\n\\end{enumerate}\n"

footers_string = """
\\section{Conclusion}
\\section{Acknowledgements}
\\section{References}
References are usually optional
\\end{document}
"""


final_string = (headers
                + title_string
                + authors_string
                + date_string
                + begin_document_string
                + make_title_string
                + executive_summary_string
                + introduction_string
                + experimental_methodology_string
                + results_string
                + discussion_string
                + footers_string)

with open(change_extension(argument), 'w') as file:
    file.write(final_string)

                

