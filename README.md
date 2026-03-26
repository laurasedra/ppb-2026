# ppb-2026
Building the interface for my university's budget project.
There are many elements here that are all tied together. 


refresh_scores_view.py:
  goes with a spreadsheet that is used to collect data from forms used for voting. 

FIXME
  The only things that need to be changed to keep the integrity of the code is while also customizing to your own files:
    WORKBOOK_PATH
    
extractor_parser.py:
  extracts voting data from corresponding powerpoint and populates initials into sheet, populates a tally of all the votes.
