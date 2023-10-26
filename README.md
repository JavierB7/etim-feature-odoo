# BMEcat XLS to Odoo attributes
Read BMECat XLS file, decode Features with ETIM API and create Odoo's attributes XLS file.

## Prerequisites
- Install the dependencies:

`pip install -r requirements.txt`

- Get a client_id and client_secret from the [ETIM API](https://etimapi.etim-international.com/).

## Use
Run the script with the following arguments:

- Data file name: The XLS file name with the BMEcat data. It should be in XLS format.
- ETIM client ID: The client ID for the ETIM API
- ETIM client secret: The client secret for the ETIM API

Example:

`python3 etim.py data_file.xls my_etim_client_id my_etim_client_secret`
