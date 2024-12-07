# yourcareer-scripts


The main purpose of the scripts is to generate word documents formatted correctly for badges (for either students or company people). The formats for the students and companies are different, hence the two different word files.

The excel files to be used for badges are the ones generated by JobTeaser. The excel file for company people should be formatted as `"First name", "Last name", "Company name"`
- ``generate_badges.py`` requires all the excel files pertaining to the students to end with "students.csv". These files should be: speed dates or events such as the lunch or photoshoot. 
- ``generate_badges_company_people.py`` requires the excel file pertaining to the companies to be named "company_people.csv".
- ``generate_attendance_lists_indiv.py`` should be put in a separate folder where all the excel sheets with the speed dates should be present.

To run the scripts, you need to have [Python](https://www.python.org/downloads) and [python-docx](https://python-docx.readthedocs.io/en/latest/) installed.
