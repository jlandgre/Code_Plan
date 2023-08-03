This repository contains Python utilities for converting code to a rows/columns "Code Plan" spreadsheet table. The table is useful for planning projects and for automating creation of code from the plan --either with utilities in CodePlan.py or with AI tools.

Python Classes/Use Cases in CodePlan.py
* VBAToCodePlan - Create Code Plan spreadsheet table from VBA subroutines and functions in code modules imported as text
* CodePlanToVBA - [Planned as of August 2023] Create VBA function/sub and test skeleton code and pasteable instructions for code writing by AI tools
* CodePlanToPython - [Planned as of August 2023] Create Python function and test skeleton code and pasteable instructions for code writing by AI tools
* PythonToCodePlan - Create Code Plan spreadsheet table from Python functions in *.py files

AI tools such as GPT 3.5 and GPT 4 and Github copilot are useful but are difficult to instruct for repeatable mass creation of code comprised of multiple single-action functions. The AI tools are also non-ideal for converting code from one language to another such as from VBA to Python. It seems easier to do this programmatically with Python utilities such as in this repository. A structured Code Plan spreadsheet makes a waystation for listing code details for conversion and for making use of AI tools as much as possible. The Code Plan spreadsheet facilitates automatic creation of Python functions and Pytest tests "skeleton" code or creation of code skeletons for VBA functions and associated VBA test code. 

<p align="center">
  Example Code Plan Spreadsheet</br>
  <img src=images/code_plan_spreadsheet.png "Example Code Plan Spreadsheet" width=600></br>
</p>

While not yet implemented, future CodePlan.py scripts can extract code internals from multiple functions and include the internals in the Code Plan spreadsheet. This formats the internals in a format suitable for AI language conversion. 

J.D. landgrebe