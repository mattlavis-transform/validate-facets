# Validate facets used in search

## Installation

- Create and activate a virtual environment, e.g.

  `python3 -m venv venv/`
  `source venv/bin/activate`

- Set up local variables

  - `FACETS_SOURCE` = the location of the source Excel file
  - `CLASSIFIER_FOLDER` = where the classifiers are stored

- Install necessary Python modules 

  via `pip3 install -r requirements.txt`
  
- To run

  `python3 validate.py`


- Test remotely

  `https://comm-code-history.herokuapp.com/`


- To restart the front-end and back-end services
  `./startup.sh`
