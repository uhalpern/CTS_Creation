# CTS_Creation

A brief description of the project, its purpose, and key features.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Installation

Follow the steps below to set up the project on your local machine:

### Prerequisites
Make sure you have the following installed:
- [Python 3.12+](https://www.python.org/downloads/)
- [pipenv](https://pipenv.pypa.io/en/latest/install/)

To install `pipenv`, run:
```bash
pip install pipenv
```

### Setting Up Project

1. Clone the repository:
    ```bash
    git clone https://github.com/uhalpern/CTS_Creation.git
    cd CTS_Creation
    ```
2. Download project dependencies using `pipenv` with:
    ```bash
    pipenv install
    ```

    This will automatically install the packages listed in `Pipfile` and `Pipfile.lock`

3. (Optional) - Install Developer Packages:
    ```bash
    pipenv install --dev
    ```

## Usage

### Generating Example Spreadsheet

The configuration and fake data is already created for the example spreadsheet. Just run the following:

1. Activate Environment:

    ```bash
    pipenv shell
    ```
2. Remove `python_CTS_example.xlsx` from `generated_sheets`
3. Run the main script
    ```bash
    python main.py
    ```


