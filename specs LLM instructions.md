## Goal: prototype a quick GUI with .bat for the following requirements

Use the Gemini .env we have from other folders in the internal apps

use the Claude .env we have from other folders in the internal apps

.bat to launch the GUI on any PC (make sure requirements are asked or installed if not presented on the user machine, do not use virtual enviroment)
Include requirements.txt
Include a readme.md
A basic guideline.md

Implement the following as a prototype with TKINTER/GUI:

## SKILLS:

- API USAGE (agentic workflows)
- LOC ENGINEERING
- DEVELOPMENT
- SOLUTIONS

##SPECIFICATIONS

File formats: docx

INPUT: 1 article (testing)
REFERENCE 1: reference materials
Last one (important): Internal Spot Check or Any number

Output:

1) Article copied with highlight and automatic comments based on:

Reference Materials
Internal Spot Check

2) Summarise of all issues:

Article name
Summary
breakline
next article

3) Generate a report/score in excel format with assesment

QA/Readiness per article as an excel/xlsx

-------------------------------------------------------------------
Workflow - Agentic Workflow 2 steps:


Upload stage:

User uploads folder with article docxs (any file)
User uploads reference materials folder
user uploads internal spot check

We extract docx using Pythob library as JSON files -> this is flat JSON content or markdown
--------------------------------------------------------------------------------------------
We pass the JSON or Markdown to the first Agent via API calls:

Agent 1 (API): proceed (gemini) - First agent summarise all the markdowns/JSONs and generate the QA/Spotchecks dataframe

From the refence materials and internal spotchecks -> we generate a list of QA/SpotChecks references points ie: Uppercase not allowed in Branding Name
Once the spotchecks/qa reference are generated

Agent 2 (API): proceed (Claude) - Second agent proceed to compare the summarised Markdown/JSONs against each article and generate articles commented + summarise docx unified + the QA readiness excel report

Proceed to compare the Article against the reference QA/SPOTCHECKS generated list

Proceed to generate the copy of the article with highlitghs and comments
Proceed to generate a summary as per step number 2 output showing all the articles
Proceed to generate one QA/Readiness from all articles listing individually and an overall score QA/readiness + generate an extra sheet called WordCount listing wordcount for all files uploaded article, per reference materials & internal spotcheck -> 20 files = 400k words = 2k USD (gemini/claude) (0.005 usd)
