# PyCor

## What is this?
This project is for automated homework/exercise marking, based on 
Excel sheets.

It's achieved by checking an email inbox at a configured interval, comparing
the file names of the submitted files to known code names[0], writing some
submitted values and a student's matriculation number[1] into the 
`corrector.xlsx` and then comparing the generated values to the student's 
values.

This project was developed as part of my work at [FH Aachen], after taking over 
maintenance of the earlier version by [@davahue] and [@FlowCV].

[0] configured in a `corrector.xlsx`/`.xlsm` at B2  
[1] may be used to generate different solutions per student

## Deployment

### Prerequisites
Since this project is based on Excel sheets and Windows's COM object interfaces,
you will need an installation of Windows and Excel.

- [Python 3.7](https://www.python.org/)
- [pipenv](https://pipenv.pypa.io/en/latest/)
- [pip](https://pip.pypa.io/en/stable/)

`pip` is part of the default Python installation on Windows, `pipenv` can be
installed by running:
```bash
$ pip install pipenv
```

### Installing
`pipenv` will take care of installing all needed Python dependencies, you just 
need to run the following.
```bash
$ pipenv sync
```

Afterwards copy/rename `config.example.py` to `config.py` and change values 
where necessary (a later rewrite might switch to `.env` via `python-dotenv`).
```bash
$ cp pycor/config.example.py pycor/config.py
```

The script can then be run via the following command and will check for new 
mails as defined by `DELAY_SLEEP` and compare the file names to the code names 
as defined in `corrector.xlsx`.


### Usage
Pycor can then be run via `pipenv run python -m pycor` or simply 
`python -m pycor` after activating the virtual environment.

## Contributors
- Daniel B. Bung ([@FlowCV]) - Initiator of the project
- Daniel Valero ([@davahue]) - Maintainer until 2018
- Yannick Linke ([@invisi]) - Maintainer since 2018


[@davahue]: https://github.com/davahue
[@FlowCV]: https://github.com/FlowCV
[@invisi]: https://github.com/invisi
[FH Aachen]: https://www.fh-aachen.de/