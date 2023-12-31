# MITRE ATTACK UC COVERAGE

Most enterprises take follow-up of your use cases in excel file to try obtain dashboard, status, metrics and a lot of more attributes that are importants to them but this way doesn't give a picture visibility that what is covering the use cases within a kill chain. This project consist in a Python script that let you have a visibility of your UC's coverage in MITRE ATTACK Navigator.

# Requirements
1. Download python file UC_Excel_MitreAttackNavigator.py
2. Install Python 3
3. Excel with your Use Cases and each UC has to include two columns: "Tactic", "TechniqueID". For example:

|Use Case| Tactic        | TechniqueID |
|---| ------------- |-------------| 
|Detect legitimate web compromised| defense-evasion | T1006| 
|Detect public app compromised | execution | T1609| 

# Execution
1. Run UC_Excel_MitreAttackNavigator.py
2. Answer the following questions:

| Questions        | Means |
| ------------- |-------------| 
| Enter filename | Give fullname path of excel file | 
| Your excel has more than one sheet? Yes or No | Answer 'Yes' if file has more than one sheet | 
| What is the name Excel sheet? | Give name of the sheet that contains UCs, Tactic, Technique | 
| What is the excel column that store tactics? | Give letter of column that contain tactics in excel | 
| What is the column that store techniques? | Give letter of column that contain techniques in excel | 

4. If all was executed sucessfull, script'll create a json file "MitreNavigator.json".
5. Upload json file "MitreNavigator.json" in https://mitre-attack.github.io/attack-navigator/ to see what tactics and techniques are covering wiht your use cases.

** Note: It was tested with MITRE ATT&CK® Matrix for Enterprise. Not sure if it works for Mobile Matrix and/or ICS Matrix

# Sample

![Sample](/Screen_Sample.png?raw=true "Sample")
![Sample](/MitreNavigatorHeatMap.png?raw=true "Sample")
