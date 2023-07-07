# Power Test Campaign Generator

The Power Test Campaign Generator is a Python script that automates the process of generating Power Test Campaigns (TC) for power-related testing. It takes input from Excel files containing test data and creates test cases and campaigns accordingly.

## Features

- Generates Power TC campaigns based on Excel files and test data.
- Provides a command-line interface for easy interaction.
- Supports multiple Excel files and sheets selection.
- Overwrites existing campaigns or creates new campaigns with unique names.
- Replaces variables in test cases based on provided test data.
- Uses default test cases and procedures as templates for generating the campaigns.

## Prerequisites

- Python 3.7 or above
- openpyxl library (`pip install openpyxl`)

## Usage

1. Clone the repository:
   ```bash
   git clone https://github.com/Alongri/PowerTCgenerator.git
2. Navigate to the project directory:
    ```bash
    cd PowerTCgenerator
3. Install the required dependencies:
   ```bash
    pip install -r requirements.txt
4. Run the script:
   ```bash
    python power_tc_generator.py

## Excel File Format
The script expects Excel files to follow a specific format to ensure proper data extraction and generation of test cases. Please make sure your Excel files adhere to the following guidelines:

- Each sheet in the Excel file represents a specific category of power tests.
- The first row of each sheet should contain the column headers.
- The first column of each sheet should contain the test case names.
- The second column of each sheet should contain the test case descriptions.
- Subsequent columns can contain test data specific to each test case.

## File Structure
- power_tc_generator.py: The main Python script for generating the Power TC campaign.
- requirements.txt: A list of required Python dependencies.
- README.md: This file, providing information about the project.

## Contributing
Contributions to the Power Test Campaign Generator are welcome! If you find any bugs, have suggestions for improvements, or would like to add new features, please feel free to open an issue or submit a pull request. Your contributions can help make this project even better.

## License

This project is licensed under the [MIT License](LICENSE).

## Author
Alon Gritsovsky
